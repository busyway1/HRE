VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDirectory 
   Caption         =   "����(����) ����"
   ClientHeight    =   9900.001
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   13068
   OleObjectBlob   =   "frmDirectory.frx":0000
   StartUpPosition =   1  '������ ���
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
    
    On Error Resume Next
    
    If Me.File_Option = False And Me.Folder_Option = False Then
        Msg "����(����) �ɼ��� üũ���ּ���!", vbCritical
        Exit Sub
    End If
    
    Set tbl = HideSheet.ListObjects("TempPath")
    Set dataRange = tbl.DataBodyRange
    Set foundRow = dataRange.Find(What:=Me.Name_Combo.Value, _
                                  LookIn:=xlValues, _
                                  LookAt:=xlWhole, _
                                  SearchOrder:=xlByRows, _
                                  SearchDirection:=xlNext, _
                                  MatchCase:=False, _
                                  SearchFormat:=False)
    If Not foundRow Is Nothing Then
        If Me.File_Option = True Then
            If foundRow.Cells(1, tbl.ListColumns("����").Index - 1).Value = "����" Then
                ' ���� Ȯ���� �˻�
                If Not HasValidFileExtension(Me.Path_Text.Value) Then
                    userResponse = MsgBox("�Է��� ������ �����̰ų� �����Ǵ� Ȯ���ڰ� �ƴմϴ�. ����Ͻðڽ��ϱ�?", vbYesNo + vbQuestion, AppName & " " & AppType)
                    If userResponse = vbNo Then Exit Sub
                End If
                foundRow.Cells(1, tbl.ListColumns("���").Index - 1).Value = Me.Path_Text.Value
            Else
                Msg "������ �׸��� '����'�� �ƴմϴ�.", vbExclamation
                Exit Sub
            End If
        ElseIf Me.Folder_Option = True Then
            If foundRow.Cells(1, tbl.ListColumns("����").Index - 1).Value = "����" Then
                ' ���� ��� �˻�
                If HasValidFileExtension(Me.Path_Text.Value) Then
                    userResponse = MsgBox("�Է��� ��ΰ� ����ó�� ���Դϴ�. ������ ����Ͻðڽ��ϱ�?", vbYesNo + vbQuestion, AppName & " " & AppType)
                    If userResponse = vbNo Then Exit Sub
                End If
                foundRow.Cells(1, tbl.ListColumns("���").Index - 1).Value = Me.Path_Text.Value
            Else
                Msg "������ �׸��� '����'�� �ƴմϴ�.", vbExclamation
                Exit Sub
            End If
        End If
        

        With Check.Cells(13, 4)
            .Value = "Complete"
            .Interior.Color = RGB(198, 239, 206)
            .Offset(0, 1).Value = Format(Now(), "yyyy-mm-dd hh:mm")
            .Offset(0, 2).Value = GetUserInfo()
        End With
        
        Msg "�Է� �ּҰ� ��ϵǾ����ϴ�.", vbInformation
        isNormalClose = True
        Unload Me
    Else
        Msg "������ �׸��� ã�� �� �����ϴ�.", vbExclamation
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
        GoEnd "���� �ܰ踦 �Ϸ����ּ���!"
    End If
    With Check.Cells(13, 4)
        .Value = "In Progress"
        .Interior.Color = RGB(255, 235, 156)
        .Offset(0, 1).Value = Format(Now(), "yyyy-mm-dd hh:mm")
        .Offset(0, 2).Value = GetUserInfo()
    End With
    
    Set qt = DirectoryURL.ListObjects(1).QueryTable
    Call SpeedUp
    Call OpenProgress("SPO�� ���� ��...")
    Call CalculateProgress(0.5, "SPO�κ��� ���͸� ���� ���� ��...")
    qt.Refresh BackgroundQuery:=False
    Application.CalculateUntilAsyncQueriesDone
    InitializeDirectoryStructure
    PopulateTreeView
    Call CalculateProgress(1, "�۾� �Ϸ�")
    Call SpeedDown
    Set qt = Nothing
End Sub
Private Sub InitializeDirectoryStructure()
    Dim tbl As ListObject
    Dim dataRange As Range
    Dim i As Long, j As Long
    Dim currentDir As Directory
    On Error Resume Next
    
    Set tbl = DirectoryURL.ListObjects("���͸�")
    Set dataRange = tbl.DataBodyRange
    
    Set rootDirectory = New Directory
    rootDirectory.path = HideSheet.Range("E2").Value
    Set rootDirectory.SubDirectories = New Collection
    
    For i = 1 To dataRange.Rows.count
        Set currentDir = rootDirectory
        For j = 1 To dataRange.Columns.count
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

    uniqueKey = Replace(uniqueKey, " ", "_")
    uniqueKey = Replace(uniqueKey, "\", "_")
    uniqueKey = Replace(uniqueKey, "/", "_")
    uniqueKey = Replace(uniqueKey, ":", "_")
    uniqueKey = Replace(uniqueKey, ".", "_")
    uniqueKey = Replace(uniqueKey, "&", "_")
    uniqueKey = Replace(uniqueKey, "%", "_")
    uniqueKey = Replace(uniqueKey, "=", "_")
    uniqueKey = uniqueKey & "_" & CStr(TreeView.Nodes.count + 1)
    
    If Len(Trim(uniqueKey)) = 0 Or Len(uniqueKey) > 255 Then
        uniqueKey = "default_key_" & TreeView.Nodes.count + 1
    End If

    ' TreeView���� ������ ����� ��尡 �ִ��� Ȯ��
    For Each existingNode In TreeView.Nodes
        If existingNode.key = uniqueKey Then
            nodeExists = True
            Exit For
        End If
    Next existingNode

    ' �ߺ��� ��尡 ���� ��쿡�� �߰�
    If Not nodeExists Then
        If parentNode Is Nothing Then
            ' ��Ʈ ��� �߰�
            Set newNode = TreeView.Nodes.Add(, , uniqueKey, dir.path)
        Else
            ' ���� ��� �߰�
            Set newNode = TreeView.Nodes.Add(parentNode, tvwChild, uniqueKey, dir.path)
        End If
    Else
        ' �̹� �����ϴ� ��� ���� ��带 ���
        Set newNode = existingNode
    End If

    ' ���� ���͸� �߰�
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
        GetFullPath = Node.text
    Else
        GetFullPath = GetFullPath(Node.Parent) & "/" & Node.text
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
    
    For Each cell In dataRange.Columns(tbl.ListColumns("����").Index).Cells
        relativeRow = cell.row - dataRange.row + 1
        If (Me.File_Option And cell.Value = "����") Or (Me.Folder_Option And cell.Value = "����") Then
            itemList = itemList & ";" & dataRange.Cells(relativeRow, tbl.ListColumns("�̸�").Index).Value
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
    Set dataRange = tbl.ListColumns("�̸�").DataBodyRange
    Set foundCell = dataRange.Find(What:=selectedValue, LookIn:=xlValues, LookAt:=xlWhole)
    
    If Not foundCell Is Nothing Then
        Dim descriptionValue As String
        Dim presentPath As String
        descriptionValue = tbl.ListColumns("Description").DataBodyRange.Cells(foundCell.row - dataRange.row + 1, 1).Value
        presentPath = tbl.ListColumns("���").DataBodyRange.Cells(foundCell.row - dataRange.row + 1, 1).Value
        Me.Description_Text.Value = descriptionValue
        Me.PresentPath_Text.Value = presentPath
    Else
        Me.Description_Text.Value = "�ش� ����(����) ���� ������ �������� �ʽ��ϴ�."
        Me.PresentPath_Text.Value = "�ش� ����(����) ���� ��ΰ� �������� �ʽ��ϴ�."
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
