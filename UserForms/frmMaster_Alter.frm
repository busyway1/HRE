VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMaster_Alter 
   Caption         =   "Master 수정"
   ClientHeight    =   4995
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   7512
   OleObjectBlob   =   "frmMaster_Alter.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "frmMaster_Alter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mRowIndex As Long
Public mColIndex As Long
Private logVal As String
Public Property Let colIndex(Value As Long)
    mColIndex = Value
End Property
Public Property Get colIndex() As Long
    colIndex = mColIndex
End Property
Public Property Let rowIndex(Value As Long)
    mRowIndex = Value
End Property
Public Property Get rowIndex() As Long
    rowIndex = mRowIndex
End Property
Public Sub LoadData(cellarray() As Variant, rowIndex As Long, colIndex As Long)
    On Error Resume Next
    Dim i As Long
    Dim tbl As ListObject
    
    Set tbl = CoAMaster.ListObjects("Master")
    
    Me.SelectCol_Text.Value = tbl.HeaderRowRange(colIndex).Value
    Me.PwCAccCode_Text.Value = cellarray(1).Value
    Me.PwCAccName_Text.Value = cellarray(2).Value
    Me.SelectColVal_Text.Value = tbl.ListRows(rowIndex).Range(colIndex).Value
    
    logVal = Me.SelectColVal_Text.Value
    Me.rowIndex = rowIndex
    Me.colIndex = colIndex
    
    Set tbl = Nothing
End Sub
Private Sub Alter_Command_Click()
    On Error Resume Next
    Dim tbl As ListObject
    Set tbl = CoAMaster.ListObjects("Master")
    
    CoAMaster.Unprotect PASSWORD
    
    tbl.ListRows(Me.rowIndex).Range(Me.colIndex).Value = Me.SelectColVal_Text.Value
    With Check.Cells(17, 4)
        .Value = "If Any"
        .Interior.Color = RGB(237, 237, 237)
        .Offset(0, 1).Value = Format(Now(), "yyyy-mm-dd hh:mm")
        .Offset(0, 2).Value = GetUserInfo()
    End With

    '로그 전송
    LogData CoAMaster.name, "<Master 변경>" & vbNewLine & vbNewLine & _
                     "[변경 전]" & vbNewLine & _
                     "PwC_CoA: " & Me.PwCAccCode_Text.Value & vbNewLine & _
                     "PwC_계정명: " & Me.PwCAccName_Text.Value & vbNewLine & _
                     "선택 열: " & Me.SelectCol_Text.Value & vbNewLine & _
                     "열값: " & logVal & vbNewLine & vbNewLine & _
                     "[변경 후]" & vbNewLine & _
                     "PwC_CoA: " & Me.PwCAccCode_Text.Value & vbNewLine & _
                     "PwC_계정명: " & Me.PwCAccName_Text.Value & vbNewLine & _
                     "선택 열: " & Me.SelectCol_Text.Value & vbNewLine & _
                     "열값: " & Me.SelectColVal_Text.Value
    
    CoAMaster.Protect PASSWORD, UserInterfaceOnly:=True, AllowFiltering:=True
    
    Msg "데이터가 성공적으로 수정되었습니다.", vbInformation
    Unload Me
    
    Set tbl = Nothing
End Sub
Private Sub Cancel_Command_Click()
    On Error Resume Next
    Unload Me
End Sub
Private Sub UserForm_Initialize()
    Me.Caption = AppName & " " & AppType
End Sub
