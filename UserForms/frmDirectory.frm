VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDirectory
   Caption         =   "파일(폴더) 경로 선택"
   ClientHeight    =   9900.001
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   13068
   OleObjectBlob   =   "frmDirectory.frx":0000
   StartUpPosition =   1  '소유자 중앙
End
Attribute VB_Name = "frmDirectory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private isNormalClose As Boolean
Private rootDirectory As Directory

' Late Binding용 상수 선언 (MSComctlLib 참조 없이 사용)
Private Const tvwChild As Long = 4
Private Const tvwFirst As Long = 0
Private Const tvwLast As Long = 1
Private Const tvwNext As Long = 2
Private Const tvwPrevious As Long = 3

Private Sub Select_Cmd_Click()
    Dim tbl As ListObject
    Dim dataRange As Range
    Dim foundRow As Range
    Dim userResponse As VbMsgBoxResult
    Dim rowIndex As Long

    On Error Resume Next

    If Me.File_Option = False And Me.Folder_Option = False Then
        Msg "파일(폴더) 옵션을 체크해주세요!", vbCritical
        Exit Sub
    End If

    Set tbl = HideSheet.ListObjects("TempPath")
    Set dataRange = tbl.ListColumns("이름").DataBodyRange
    Set foundRow = dataRange.Find(What:=Me.Name_Combo.Value, _
                                  LookIn:=xlValues, _
                                  LookAt:=xlWhole, _
                                  SearchOrder:=xlByRows, _
                                  SearchDirection:=xlNext, _
                                  MatchCase:=False, _
                                  SearchFormat:=False)
    If Not foundRow Is Nothing Then
        rowIndex = foundRow.Row - tbl.DataBodyRange.Row + 1

        If Me.File_Option = True Then
            If tbl.DataBodyRange.Cells(rowIndex, tbl.ListColumns("구분").Index).Value = "파일" Then
                ' 파일 확장자 검사
                If Not HasValidFileExtension(Me.Path_Text.Value) Then
                    userResponse = MsgBox("입력한 파일이 폴더이거나 지원되는 확장자가 아닙니다. 계속하시겠습니까?", vbYesNo + vbQuestion, AppName & " " & AppType)
                    If userResponse = vbNo Then Exit Sub
                End If
                tbl.DataBodyRange.Cells(rowIndex, tbl.ListColumns("경로").Index).Value = Me.Path_Text.Value
            Else
                Msg "선택한 항목이 '파일'이 아닙니다.", vbExclamation
                Exit Sub
            End If
        ElseIf Me.Folder_Option = True Then
            If tbl.DataBodyRange.Cells(rowIndex, tbl.ListColumns("구분").Index).Value = "폴더" Then
                ' 폴더 경로 검사
                If HasValidFileExtension(Me.Path_Text.Value) Then
                    userResponse = MsgBox("입력한 경로가 파일처럼 보입니다. 폴더로 계속하시겠습니까?", vbYesNo + vbQuestion, AppName & " " & AppType)
                    If userResponse = vbNo Then Exit Sub
                End If
                tbl.DataBodyRange.Cells(rowIndex, tbl.ListColumns("경로").Index).Value = Me.Path_Text.Value
            Else
                Msg "선택한 항목이 '폴더'가 아닙니다.", vbExclamation
                Exit Sub
            End If
        End If


        With Check.Cells(13, 4)
            .Value = "Complete"
            .Interior.Color = RGB(198, 239, 206)
            .Offset(0, 1).Value = Format(Now(), "yyyy-mm-dd hh:mm")
            .Offset(0, 2).Value = GetUserInfo()
        End With

        Msg "입력 주소가 기록되었습니다.", vbInformation
        isNormalClose = True
        Unload Me
    Else
        Msg "선택한 항목을 찾을 수 없습니다.", vbExclamation
    End If

    Set tbl = Nothing: Set dataRange = Nothing: Set foundRow = Nothing
End Sub
Private Sub Cancel_Cmd_Click()
    Unload Me

    With Check.Cells(13, 4)
        .Value = "Not Started"
        .Interior.Color = RGB(255, 199, 206)
        .Offset(0, 1).Value = Format(Now(), "yyyy-mm-dd hh:mm")
        .Offset(0, 2).Value = GetUserInfo()
    End With
    GoEnd
End Sub
Private Sub UserForm_Initialize()
    Dim qt As QueryTable
    On Error Resume Next
    Me.Caption = AppName & " " & AppType
    If Check.Cells(12, 4).Value <> "Complete" Then
        GoEnd "이전 단계를 완료해주세요!"
    End If
    With Check.Cells(13, 4)
        .Value = "In Progress"
        .Interior.Color = RGB(255, 235, 156)
        .Offset(0, 1).Value = Format(Now(), "yyyy-mm-dd hh:mm")
        .Offset(0, 2).Value = GetUserInfo()
    End With

    Set qt = DirectoryURL.ListObjects(1).QueryTable
    Call SpeedUp
    Call OpenProgress("SPO에 접근 중...")
    Call CalculateProgress(0.5, "SPO로부터 디렉터리 정보 취합 중...")
    qt.Refresh BackgroundQuery:=False
    Application.CalculateUntilAsyncQueriesDone
    InitializeDirectoryStructure
    PopulateTreeView
    Call CalculateProgress(1, "작업 완료")
    Call SpeedDown
    Set qt = Nothing
End Sub
Private Sub InitializeDirectoryStructure()
    Dim tbl As ListObject
    Dim dataRange As Range
    Dim i As Long, j As Long
    Dim currentDir As Directory
    On Error Resume Next

    Set tbl = DirectoryURL.ListObjects("디렉터리")
    Set dataRange = tbl.DataBodyRange

    Set rootDirectory = New Directory
    rootDirectory.path = HideSheet.Range("E2").Value
    Set rootDirectory.SubDirectories = New Collection

    For i = 1 To dataRange.Rows.Count
        Set currentDir = rootDirectory
        For j = 1 To dataRange.Columns.Count
            Dim cellValue As String
            cellValue = dataRange.Cells(i, j).Value
            If cellValue <> "" Then
                Dim newDir As Directory
                Set newDir = FindOrCreateSubDirectory(currentDir, cellValue)
                Set currentDir = newDir
            End If
        Next j
    Next i

    Set tbl = Nothing: Set dataRange = Nothing
End Sub
Private Sub PopulateTreeView()
    TreeView.Nodes.Clear
    AddDirectoryToTreeView rootDirectory, Nothing
End Sub
Private Sub AddDirectoryToTreeView(dir As Directory, parentNode As Object)
    Dim newNode As Object
    Dim uniqueKey As String
    Dim nodeExists As Boolean
    Dim existingNode As Object
    nodeExists = False
    On Error Resume Next

    ' BUGFIX: uniqueKey를 dir.path로 초기화해야 함 (기존에는 빈 문자열이었음)
    uniqueKey = dir.path
    uniqueKey = Replace(uniqueKey, " ", "_")
    uniqueKey = Replace(uniqueKey, "\", "_")
    uniqueKey = Replace(uniqueKey, "/", "_")
    uniqueKey = Replace(uniqueKey, ":", "_")
    uniqueKey = Replace(uniqueKey, ".", "_")
    uniqueKey = Replace(uniqueKey, "&", "_")
    uniqueKey = Replace(uniqueKey, "%", "_")
    uniqueKey = Replace(uniqueKey, "=", "_")
    uniqueKey = uniqueKey & "_" & CStr(TreeView.Nodes.Count + 1)

    If Len(Trim(uniqueKey)) = 0 Or Len(uniqueKey) > 255 Then
        uniqueKey = "default_key_" & TreeView.Nodes.Count + 1
    End If

    ' TreeView에서 동일한 경로의 노드가 있는지 확인
    For Each existingNode In TreeView.Nodes
        If existingNode.Key = uniqueKey Then
            nodeExists = True
            Exit For
        End If
    Next existingNode

    ' 중복된 노드가 없는 경우에만 추가
    If Not nodeExists Then
        If parentNode Is Nothing Then
            ' 루트 노드 추가
            Set newNode = TreeView.Nodes.Add(, , uniqueKey, dir.path)
        Else
            ' 하위 노드 추가
            Set newNode = TreeView.Nodes.Add(parentNode, tvwChild, uniqueKey, dir.path)
        End If
    Else
        ' 이미 존재하는 경우 기존 노드를 사용
        Set newNode = existingNode
    End If

    ' 하위 디렉터리 추가
    Dim subDir As Directory
    For Each subDir In dir.SubDirectories
        AddDirectoryToTreeView subDir, newNode
    Next subDir
End Sub
Public Function FindOrCreateSubDirectory(parentDir As Directory, name As String) As Directory
    On Error Resume Next
    Dim subDir As Directory
    For Each subDir In parentDir.SubDirectories
        If subDir.path = name Then
            Set FindOrCreateSubDirectory = subDir
            Exit Function
        End If
    Next subDir
    Set subDir = New Directory
    subDir.path = name
    parentDir.SubDirectories.Add subDir
    Set FindOrCreateSubDirectory = subDir
End Function
Public Function FindDirectory(dir As Directory, path As String) As Directory
    On Error Resume Next
    If dir.path = path Then
        Set FindDirectory = dir
        Exit Function
    End If
    Dim subDir As Directory
    For Each subDir In dir.SubDirectories
        Dim foundDir As Directory
        Set foundDir = FindDirectory(subDir, path)
        If Not foundDir Is Nothing Then
            Set FindDirectory = foundDir
            Exit Function
        End If
    Next subDir
    Set FindDirectory = Nothing
End Function
' TreeView 클릭 이벤트 - Late Binding 방식
Private Sub TreeView_Click()
    On Error Resume Next
    Dim Node As Object
    Set Node = TreeView.SelectedItem
    If Not Node Is Nothing Then
        Path_Text.Text = GetFullPath(Node)
    End If
End Sub

Private Function GetFullPath(ByVal Node As Object) As String
    On Error Resume Next
    If Node.Parent Is Nothing Then
        GetFullPath = Node.Text
    Else
        GetFullPath = GetFullPath(Node.Parent) & "/" & Node.Text
    End If
End Function
Private Sub File_Option_Click()
    On Error Resume Next
    UpdateNameCombo
End Sub
Private Sub Folder_Option_Click()
    On Error Resume Next
    UpdateNameCombo
End Sub
Private Sub UpdateNameCombo()
    On Error Resume Next
    Dim tbl As ListObject
    Dim dataRange As Range
    Dim cell As Range
    Dim itemList As String
    Dim relativeRow As Long

    Set tbl = HideSheet.ListObjects("TempPath")
    Set dataRange = tbl.DataBodyRange
    itemList = ""

    For Each cell In dataRange.Columns(tbl.ListColumns("구분").Index).Cells
        relativeRow = cell.Row - dataRange.Row + 1
        If (Me.File_Option And cell.Value = "파일") Or (Me.Folder_Option And cell.Value = "폴더") Then
            itemList = itemList & ";" & dataRange.Cells(relativeRow, tbl.ListColumns("이름").Index).Value
        End If
    Next cell

    If Len(itemList) > 0 Then
        itemList = Mid(itemList, 2)
    End If

    Me.Name_Combo.Clear
    Me.Name_Combo.Value = ""
    Me.Name_Combo.AddItem ""

    If Len(itemList) > 0 Then
        Me.Name_Combo.List = Split(itemList, ";")
    End If

    Set tbl = Nothing: Set dataRange = Nothing: Set cell = Nothing
End Sub
Private Sub Name_Combo_Change()
    On Error Resume Next
    Dim tbl As ListObject
    Dim dataRange As Range
    Dim foundCell As Range
    Dim selectedValue As String

    selectedValue = Me.Name_Combo.Value
    If selectedValue = "" Then
        Me.Description_Text.Value = ""
        Me.PresentPath_Text.Value = ""
        Exit Sub
    End If

    Set tbl = HideSheet.ListObjects("TempPath")
    Set dataRange = tbl.ListColumns("이름").DataBodyRange
    Set foundCell = dataRange.Find(What:=selectedValue, LookIn:=xlValues, LookAt:=xlWhole)

    If Not foundCell Is Nothing Then
        Dim descriptionValue As String
        Dim presentPath As String
        Dim rowIndex As Long
        rowIndex = foundCell.Row - dataRange.Row + 1
        descriptionValue = tbl.ListColumns("Description").DataBodyRange.Cells(rowIndex, 1).Value
        presentPath = tbl.ListColumns("경로").DataBodyRange.Cells(rowIndex, 1).Value
        Me.Description_Text.Value = descriptionValue
        Me.PresentPath_Text.Value = presentPath
    Else
        Me.Description_Text.Value = "해당 파일(폴더) 관련 설명이 존재하지 않습니다."
        Me.PresentPath_Text.Value = "해당 파일(폴더) 관련 경로가 존재하지 않습니다."
    End If

    Set tbl = Nothing: Set dataRange = Nothing: Set foundCell = Nothing
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    On Error Resume Next
    If Not isNormalClose Then
        With Check.Cells(13, 4)
            .Value = "Not Started"
            .Interior.Color = RGB(255, 199, 206)
            .Offset(0, 1).Value = Format(Now(), "yyyy-mm-dd hh:mm")
            .Offset(0, 2).Value = GetUserInfo()
        End With
        GoEnd
    End If
End Sub
Private Function HasValidFileExtension(path As String) As Boolean
    On Error Resume Next
    Dim validExtensions As Variant
    Dim ext As String
    Dim i As Long
    validExtensions = Array(".xlsx", ".xlsm", ".xls", ".pdf", ".doc", ".docx", ".ppt", ".pptx")
    ext = LCase(Right(path, Len(path) - InStrRev(path, ".")))
    For i = LBound(validExtensions) To UBound(validExtensions)
        If ext = LCase(Mid(validExtensions(i), 2)) Then
            HasValidFileExtension = True
            Exit Function
        End If
    Next i
    HasValidFileExtension = False
End Function
