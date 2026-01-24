VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCoA_Update 
   Caption         =   "CoA 추가"
   ClientHeight    =   6900
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   7548
   OleObjectBlob   =   "frmCoA_Update.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "frmCoA_Update"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private whichWS As Worksheet '어떤 시트에서 유저폼을 켰는지
Public mRowIndex As Long
Public Property Let rowIndex(Value As Long)
    mRowIndex = Value
End Property
Public Property Get rowIndex() As Long
    rowIndex = mRowIndex
End Property
Public Sub LoadData(cellarray() As Variant, rowIndex As Long)
    CorpCode_Text.Value = cellarray(1)
    CorpAccCode_Text.Value = cellarray(2)
    CorpAccName_Text.Value = cellarray(3)
    Me.rowIndex = rowIndex
End Sub
Private Sub Append_Command_Click() ' 저장 버튼
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim tblCoA As ListObject
    Dim newRow As ListRow
    Dim existingRow As ListRow
    Dim isDuplicate As Boolean
    
    On Error Resume Next
    
    Set ws = whichWS
    Set tblCoA = CorpCoA.ListObjects("Raw_CoA")

    If Me.PwCAccCode_Text.Value = "" Then
        Msg "PwC_계정명을 확인해주세요!", vbCritical
        Exit Sub
    End If
    
    Select Case whichWS.CodeName
        Case "BSPL"
            Set tbl = ws.ListObjects("PTB")
        Case "ADBS"
            Set tbl = ws.ListObjects("AD_BS")
        Case "MCCoA"
            Set tbl = ws.ListObjects("제조원가명세서")
        Case "MCCoA_AD"
            Set tbl = ws.ListObjects("제조원가명세서_취득_처분")
    End Select
    
    isDuplicate = False
    For Each existingRow In tblCoA.ListRows
        If Trim(CStr(existingRow.Range(1).Value)) = Trim(CStr(CorpCode_Text.Value)) And _
           Trim(CStr(existingRow.Range(2).Value)) = Trim(CStr(CorpAccCode_Text.Value)) Then
            isDuplicate = True
            Exit For
        End If
    Next existingRow
    
    If isDuplicate Then
        Msg "이미 존재하는 데이터입니다. 회사코드와 회사계정코드를 Raw_CoA에서 확인해주세요.", vbExclamation
        Exit Sub
    End If
    
    
    ws.Unprotect PASSWORD
    CorpCoA.Unprotect PASSWORD
    
    Set newRow = tblCoA.ListRows.Add
    newRow.Range(1) = CorpCode_Text.Value
    newRow.Range(2) = CorpAccCode_Text.Value
    newRow.Range(3) = CorpAccName_Text.Value
    newRow.Range(4) = PwCAccCode_Text.Value
    newRow.Range(5) = PwCAccName_Combo.Value
    newRow.Range(6) = Detail_Text.Value
    
    tbl.ListRows(Me.rowIndex).Range.Interior.Color = RGB(0, 176, 80)
    

    With Check.Cells(19, 4)
        .Value = "If Any"
        .Interior.Color = RGB(237, 237, 237)
        .Offset(0, 1).Value = Format(Now(), "yyyy-mm-dd hh:mm")
        .Offset(0, 2).Value = GetUserInfo()
    End With
    
    
    ws.Protect PASSWORD, UserInterfaceOnly:=True, AllowFiltering:=True
    CorpCoA.Protect PASSWORD, UserInterfaceOnly:=True, AllowFiltering:=True
    
    '로그 전송
    LogData ws.name, "<CoA 추가>" & vbNewLine & vbNewLine & _
                     "[추가 전]" & vbNewLine & _
                     "법인코드: " & vbNewLine & _
                     "계정코드: " & vbNewLine & _
                     "계정과목명: " & vbNewLine & _
                     "PwC_CoA: " & vbNewLine & _
                     "PwC_계정명: " & vbNewLine & _
                     "비고: " & vbNewLine & vbNewLine & _
                     "[추가 후]" & vbNewLine & _
                     "법인코드: " & Me.CorpCode_Text.Value & vbNewLine & _
                     "계정코드: " & Me.CorpAccCode_Text.Value & vbNewLine & _
                     "계정과목명: " & Me.CorpAccName_Text.Value & vbNewLine & _
                     "PwC_CoA: " & Me.PwCAccCode_Text.Value & vbNewLine & _
                     "PwC_계정명: " & Me.PwCAccName_Combo.Value & vbNewLine & _
                     "비고: " & Me.Detail_Text.Value
    
    MsgBox "CoA에 성공적으로 추가되었습니다.", vbInformation
    Unload Me
    
    Set ws = Nothing: Set tblCoA = Nothing: Set existingRow = Nothing
End Sub
Private Sub Cancel_Command_Click()
    Unload Me
End Sub
Private Sub UserForm_Initialize()
    Dim tbl As ListObject
    Set whichWS = ActiveSheet '어느 시트에서 활성화하였는가?
    Me.Caption = AppName & " " & AppType
    Set tbl = CoAMaster.ListObjects("Master")
    Me.PwCAccName_Combo.RowSource = "=Master[Account Name]"
    Set tbl = Nothing
End Sub
Private Sub PwCAccName_Combo_Change()
    Dim tbl As ListObject
    Dim findRange As Range
    Dim findRow As ListRow
    
    On Error Resume Next
    
    Set tbl = CoAMaster.ListObjects("Master")
    Set findRange = tbl.ListColumns("Account Name").DataBodyRange.Find(What:=CStr(Me.PwCAccName_Combo.Value), LookIn:=xlValues, LookAt:=xlWhole)
    
    If Not findRange Is Nothing Then
        Set findRow = tbl.ListRows(findRange.row - tbl.HeaderRowRange.row)
        Me.PwCAccCode_Text.Value = findRow.Range(tbl.ListColumns("TB Account").Index)
    Else
        Me.PwCAccCode_Text.Value = ""
    End If
    
    Set tbl = Nothing: Set findRange = Nothing: Set findRow = Nothing
End Sub
Private Sub PwCAccName_Combo_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    HookListBoxScroll Me, Me.PwCAccName_Combo
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    UnhookListBoxScroll
End Sub
