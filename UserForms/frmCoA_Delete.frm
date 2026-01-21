VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCoA_Delete 
   Caption         =   "CoA 삭제"
   ClientHeight    =   6900
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   7536
   OleObjectBlob   =   "frmCoA_Delete.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "frmCoA_Delete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mRowIndex As Long
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
    Me.PwCAccName_Text.Value = cellarray(5).Value
    Me.Detail_Text.Value = cellarray(6).Value
    Me.rowIndex = rowIndex
End Sub
Private Sub Delete_Command_Click()
    Dim tbl As ListObject
    Dim editRow As ListRow
    Dim response As VbMsgBoxResult
    
    On Error Resume Next
    
    Set tbl = CorpCoA.ListObjects("Raw_CoA")
    Set editRow = tbl.ListRows(Me.rowIndex)
    
    response = MsgBox("정말로 이 항목을 삭제하시겠습니까?" & vbNewLine & _
                      "법인코드: " & Me.CorpCode_Text.Value & vbNewLine & _
                      "계정코드: " & Me.CorpAccCode_Text.Value & vbNewLine & _
                      "계정과목명: " & Me.CorpAccName_Text.Value & vbNewLine & _
                      "PwC 계정코드: " & Me.PwCAccCode_Text.Value & vbNewLine & _
                      "PwC 계정명: " & Me.PwCAccName_Text.Value, _
                      vbYesNo + vbQuestion, , AppName & " " & AppType)
    
    If response = vbYes Then
         CorpCoA.Unprotect PASSWORD
         
         With Check.Cells(19, 4)
            .Value = "If Any"
            .Interior.Color = RGB(237, 237, 237)
            .Offset(0, 1).Value = Format(Now(), "yyyy-mm-dd hh:mm")
            .Offset(0, 2).Value = GetUserInfo()
        End With
         
         
         '로그 전송
        LogData CorpCoA.name, "<CoA 삭제>" & vbNewLine & vbNewLine & _
                         "[삭제 전]" & vbNewLine & _
                         "법인코드: " & Me.CorpCode_Text.Value & vbNewLine & _
                         "계정코드: " & Me.CorpAccCode_Text.Value & vbNewLine & _
                         "계정과목명: " & Me.CorpAccName_Text.Value & vbNewLine & _
                         "PwC_CoA: " & Me.PwCAccCode_Text.Value & vbNewLine & _
                         "PwC_계정명: " & Me.PwCAccName_Text.Value & vbNewLine & _
                         "비고: " & Me.Detail_Text.Value & vbNewLine & vbNewLine & _
                         "[삭제 후]" & vbNewLine & _
                         "법인코드: " & vbNewLine & _
                         "계정코드: " & vbNewLine & _
                         "계정과목명: " & vbNewLine & _
                         "PwC_CoA: " & vbNewLine & _
                         "PwC_계정명: " & vbNewLine & _
                         "비고: "
         
         editRow.Delete
         CorpCoA.Protect PASSWORD, UserInterfaceOnly:=True, AllowFiltering:=True
         
         Msg "데이터가 Raw_CoA 테이블에서 삭제되었습니다.", vbInformation
         Unload Me
     End If
    Set tbl = Nothing: Set editRow = Nothing
End Sub
Private Sub Cancel_Command_Click()
    Unload Me
End Sub
Private Sub UserForm_Initialize()
    Me.Caption = AppName & " " & AppType
End Sub
