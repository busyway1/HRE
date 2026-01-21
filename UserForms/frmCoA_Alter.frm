VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCoA_Alter 
   Caption         =   "CoA 수정"
   ClientHeight    =   6960
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   7548
   OleObjectBlob   =   "frmCoA_Alter.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "frmCoA_Alter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mRowIndex As Long
Private logVal1, logVal2, logVal3 As String
Public Property Let rowIndex(Value As Long)
    mRowIndex = Value
End Property
Public Property Get rowIndex() As Long
    rowIndex = mRowIndex
End Property
Public Sub LoadData(cellarray() As Variant, rowIndex As Long)
    Me.CorpCode_Text.Value = cellarray(1).Value
    Me.CorpAccCode_Text.Value = cellarray(2).Value
    Me.CorpAccName_Text.Value = cellarray(3).Value
    Me.PwCAccCode_Text.Value = cellarray(4).Value
    Me.PwCAccName_Combo.Value = cellarray(5).Value
    Me.Detail_Text.Value = cellarray(6).Value
    Me.rowIndex = rowIndex
    
    logVal1 = Me.PwCAccCode_Text.Value
    logVal2 = Me.PwCAccName_Combo.Value
    logVal3 = Me.Detail_Text.Value
End Sub
Private Sub Alter_Command_Click()
    Dim tbl As ListObject
    Dim editRow As ListRow
    
    On Error Resume Next

    Set tbl = CorpCoA.ListObjects("Raw_CoA")
    Set editRow = tbl.ListRows(Me.rowIndex)
    
    If Me.PwCAccCode_Text.Value = "" Then
        Msg "PwC_계정명을 확인해주세요!"
        Exit Sub
    End If
    
    CorpCoA.Unprotect PASSWORD
    
    editRow.Range(1) = CorpCode_Text.Value
    editRow.Range(2) = CorpAccCode_Text.Value
    editRow.Range(3) = CorpAccName_Text.Value
    editRow.Range(4) = PwCAccCode_Text.Value
    editRow.Range(5) = PwCAccName_Combo.Value
    editRow.Range(6) = Detail_Text.Value
    
    CorpCoA.Protect PASSWORD, UserInterfaceOnly:=True, AllowFiltering:=True
    
    With Check.Cells(19, 4)
        .Value = "If Any"
        .Interior.Color = RGB(237, 237, 237)
        .Offset(0, 1).Value = Format(Now(), "yyyy-mm-dd hh:mm")
        .Offset(0, 2).Value = GetUserInfo()
    End With
    
    '로그 전송
    LogData CorpCoA.name, "<CoA 변경>" & vbNewLine & vbNewLine & _
                     "[변경 전]" & vbNewLine & _
                     "법인코드: " & Me.CorpCode_Text.Value & vbNewLine & _
                     "계정코드: " & Me.CorpAccCode_Text.Value & vbNewLine & _
                     "계정과목명: " & Me.CorpAccName_Text.Value & vbNewLine & _
                     "PwC_CoA: " & logVal1 & vbNewLine & _
                     "PwC_계정명: " & logVal2 & vbNewLine & _
                     "비고: " & logVal3 & vbNewLine & vbNewLine & _
                     "[변경 후]" & vbNewLine & _
                     "법인코드: " & Me.CorpCode_Text.Value & vbNewLine & _
                     "계정코드: " & Me.CorpAccCode_Text.Value & vbNewLine & _
                     "계정과목명: " & Me.CorpAccName_Text.Value & vbNewLine & _
                     "PwC_CoA: " & Me.PwCAccCode_Text.Value & vbNewLine & _
                     "PwC_계정명: " & Me.PwCAccName_Combo.Value & vbNewLine & _
                     "비고: " & Me.Detail_Text.Value

    Msg "데이터가 Raw_CoA 테이블에서 수정되었습니다.", vbInformation
    Unload Me
    
    Set tbl = Nothing: Set editRow = Nothing
End Sub
Private Sub Cancel_Command_Click()
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    Dim tbl As ListObject
    Me.Caption = AppName & " " & AppType
    Set tbl = CoAMaster.ListObjects("Master")
    Me.PwCAccName_Combo.RowSource = "=Master[Account Name]"
    Set tbl = Nothing
End Sub
Private Sub PwCAccName_Combo_Change()
    Dim tbl As ListObject
    Dim findRange As Range
    Dim findRow As ListRow
    
    Set tbl = CoAMaster.ListObjects("Master")
    Set findRange = tbl.ListColumns("Account Name").DataBodyRange.Find(What:=CStr(Me.PwCAccName_Combo.Value), LookIn:=xlValues, LookAt:=xlWhole)
    
    If Not findRange Is Nothing Then
        Set findRow = tbl.ListRows(findRange.row - tbl.HeaderRowRange.row)
        Me.PwCAccCode_Text.Value = findRow.Range(tbl.ListColumns("TB Account").Index)
    Else
        Me.PwCAccCode_Text.Value = ""
    End If
    
    Set tbl = Nothing: Set findRow = Nothing
End Sub
Private Sub PwCAccName_Combo_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    HookListBoxScroll Me, Me.PwCAccName_Combo
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    UnhookListBoxScroll
End Sub
