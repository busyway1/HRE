VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmFilter 
   Caption         =   "시트 필터링"
   ClientHeight    =   4050
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   6456
   OleObjectBlob   =   "frmFilter.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "frmFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Filter_Command_Click()
    On Error Resume Next
    FilterTable
End Sub
Private Sub Cancel_Command_Click()
    On Error Resume Next
    Unload Me
End Sub
Private Sub HeaderName_Combo_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    On Error Resume Next
    HookListBoxScroll Me, Me.HeaderName_Combo
End Sub
Private Sub UserForm_Initialize()
    Me.Caption = AppName & " " & AppType
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    On Error Resume Next
    UnhookListBoxScroll
End Sub
