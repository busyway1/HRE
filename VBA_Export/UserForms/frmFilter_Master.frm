VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmFilter_Master 
   Caption         =   "시트 필터링"
   ClientHeight    =   8715.001
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   9108.001
   OleObjectBlob   =   "frmFilter_Master.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "frmFilter_Master"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private originalItems As Collection
Private selectedItems As Collection
Private Sub UserForm_Initialize()
    On Error Resume Next
    Me.Caption = AppName & " " & AppType
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim col As ListColumn
    
    Set ws = ActiveSheet
    Set tbl = ws.ListObjects(1)
    
    For Each col In tbl.ListColumns
        lstHeaders.AddItem col.name
    Next col
    lstValues.MultiSelect = fmMultiSelectMulti
    
    Set originalItems = New Collection
    Set selectedItems = New Collection
    
    ' 검색 TextBox 비활성화 (초기 상태)
    txtSearch.Enabled = False
    Set tbl = Nothing
End Sub

Private Sub lstHeaders_Click()
    On Error Resume Next
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim col As ListColumn
    Dim cell As Range
    Dim dict As Object
    Dim item As Variant
    
    Set ws = ActiveSheet
    Set tbl = ws.ListObjects(1)
    Set dict = CreateObject("Scripting.Dictionary")
    
    Set col = tbl.ListColumns(lstHeaders.Value)
    For Each cell In col.DataBodyRange
        If Not IsEmpty(cell.Value) Then
            dict(cell.Value) = 1
        End If
    Next cell
    
    lstValues.Clear
    Set originalItems = New Collection
    Set selectedItems = New Collection
 
    For Each item In SortDictionaryKeys(dict)
        lstValues.AddItem item
        originalItems.Add item
    Next item
    
    chkEmptyValues.Value = False
    txtSearch.Enabled = True
    txtSearch.Value = ""
End Sub

Private Sub txtSearch_Change()
    On Error Resume Next
    Dim keyword As String
    Dim i As Long
    Dim item As Variant
    
    keyword = LCase(Trim(txtSearch.Value))
    
    ' 현재 선택된 항목들 저장
    Set selectedItems = New Collection
    For i = 0 To lstValues.ListCount - 1
        If lstValues.Selected(i) Then
            selectedItems.Add lstValues.List(i)
        End If
    Next i
    
    lstValues.Clear
    
    For i = 1 To originalItems.count
        item = originalItems(i)
        If keyword = "" Or InStr(1, LCase(item), keyword) > 0 Then
            lstValues.AddItem item
            ' 이전에 선택되었던 항목이면 다시 선택
            If IsInCollection(selectedItems, item) Then
                lstValues.Selected(lstValues.ListCount - 1) = True
            End If
        End If
    Next i
End Sub
Private Sub txtSearch_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    On Error Resume Next
    If KeyCode = vbKeyReturn Then
        KeyCode = 0 ' 엔터 키 이벤트 취소
    End If
End Sub
Private Function IsInCollection(col As Collection, item As Variant) As Boolean
    On Error Resume Next
    Dim i As Long
    For i = 1 To col.count
        If col(i) = item Then
            IsInCollection = True
            Exit Function
        End If
    Next i
    IsInCollection = False
End Function
Private Sub cmdApplyFilter_Click()
    On Error Resume Next
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim col As ListColumn
    Dim filterCriteria As String
    Dim i As Long
    
    Set ws = ActiveSheet
    Set tbl = ws.ListObjects(1)
    Set col = tbl.ListColumns(lstHeaders.Value)
    
    ' 선택된 값들로 필터 기준 만들기
    For i = 0 To lstValues.ListCount - 1
        If lstValues.Selected(i) Then
            If filterCriteria = "" Then
                filterCriteria = "=" & lstValues.List(i)
            Else
                filterCriteria = filterCriteria & ", " & lstValues.List(i)
            End If
        End If
    Next i
    
    ' 빈 값 포함 여부 확인
    If chkEmptyValues.Value Then
        If filterCriteria = "" Then
            filterCriteria = "="
        Else
            filterCriteria = filterCriteria & ", ="
        End If
    End If
    
    ' 필터 적용
    If filterCriteria <> "" Then
        col.Range.AutoFilter Field:=col.Index, Criteria1:=Array(Split(filterCriteria, ", ")), Operator:=xlFilterValues
        MsgBox "필터링이 완료되었습니다.", vbInformation, "완료"
        Unload Me
    Else
        MsgBox "필터링할 값을 선택해주세요.", vbInformation
    End If
End Sub
Private Sub cmdClearFilter_Click()
    On Error Resume Next
    Unload Me
End Sub
Private Function SortDictionaryKeys(dict As Object) As Variant
    On Error Resume Next
    Dim keys() As Variant
    Dim i As Long, j As Long
    Dim temp As Variant
    keys = dict.keys
    For i = LBound(keys) To UBound(keys) - 1
        For j = i + 1 To UBound(keys)
            If StrComp(keys(i), keys(j), vbTextCompare) > 0 Then
                temp = keys(i)
                keys(i) = keys(j)
                keys(j) = temp
            End If
        Next j
    Next i
    SortDictionaryKeys = keys
End Function
Private Sub lstHeaders_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    On Error Resume Next
    HookListBoxScroll Me, Me.lstHeaders
End Sub
Private Sub lstValues_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    On Error Resume Next
    HookListBoxScroll Me, Me.lstValues
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    On Error Resume Next
    UnhookListBoxScroll
End Sub
