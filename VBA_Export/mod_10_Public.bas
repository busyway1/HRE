Attribute VB_Name = "mod_10_Public"
Option Explicit
' ============================================================================
' Module: mod_10_Public
' Project: HRE 연결마스터 (Consolidation Master)
' Version: 1.00
' Date: 2026-01-21
'
' Description: Global constants, utilities, and helper functions
' Adapted from BEP v1.98 for HRE consolidation requirements
' ============================================================================

' ==================== CONSTANTS ====================
Public Const PASSWORD As String = "BEP1234"
Public Const PASSWORD_Workbook As String = "PwCDA7529"
Public Const AppName As String = "HRE"
Public Const AppType = "연결마스터"
Public Const AppVersion As String = "1.00"

' ==================== GLOBAL VARIABLES ====================
Public isYellow As Long         ' CoA 확인 시 노란색으로 칠해진 행의 개수
Public isYellow_ADBS As Long    ' 취득, 처분 CoA 확인 시 노란색으로 칠해진 행의 개수

' ==================== WINDOWS API DECLARATIONS ====================
Type udtRECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

#If VBA7 Then
    Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal Index As Long) As Long
    Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
    Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
    Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal Index As Long) As Long
    Declare PtrSafe Function GetWindowRect Lib "user32" (ByVal hwnd As Long, ByRef lpRect As udtRECT) As Long
#Else
    Declare Function GetSystemMetrics Lib "user32" (ByVal Index As Long) As Long
    Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
    Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
    Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal Index As Long) As Long
    Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, ByRef lpRect As udtRECT) As Long
#End If

' ==================== ERROR HANDLING ====================
Sub GoEnd(Optional StrMsg As String)
    Call SpeedDown
    Application.AutomationSecurity = msoAutomationSecurityByUI
    If StrMsg <> vbNullString Then Msg StrMsg, vbExclamation: Call CalculateProgress(1)
    End
End Sub

Function Msg(ByVal str As String, Optional ByVal msgBoxStyle As VbMsgBoxStyle = vbOKOnly)
    Msg = MsgBox(str, msgBoxStyle, AppName & " " & AppType)
End Function

' ==================== PERFORMANCE OPTIMIZATION ====================
Sub SpeedUp()
    With Application
        .Calculation = xlCalculationManual
        .EnableEvents = False
        .ScreenUpdating = False
    End With
End Sub

Sub SpeedDown()
    With Application
        .Calculation = xlCalculationAutomatic
        .EnableEvents = True
        .ScreenUpdating = True
    End With
End Sub

' ==================== DATE FUNCTIONS ====================
Public Function RelDate() As Date ' 배포날짜
    RelDate = DateSerial(2026, 1, 21)
End Function

Public Function ExpDate() As Date ' 만료날짜
    ExpDate = DateSerial(2030, 12, 31)
End Function

Public Function CopyRight() As String
    CopyRight = "(c) " & Year(Now) & " Samil PwC. All rights reserved."
End Function

Public Function GetClosingYear() As Integer ' 결산연도
    On Error Resume Next
    GetClosingYear = CInt(HideSheet.ListObjects("결산연월").DataBodyRange(1, 1).Value)
End Function

Public Function GetClosingMonth() As Integer ' 결산월
    On Error Resume Next
    GetClosingMonth = CInt(HideSheet.ListObjects("결산연월").DataBodyRange(1, 2).Value)
End Function

' ==================== USER INFORMATION ====================
Public Function GetUserInfo() As String ' 사용자 이름 반환함수
    On Error Resume Next
    Dim userName As String
    userName = Application.userName

    ' MS Office 사용자 이름이 없으면 Windows 사용자 이름 사용
    If userName = "" Then
        userName = Environ$("USERNAME")
    End If
    GetUserInfo = userName
End Function

Public Function GetUserPath() As String  ' 파일위치 반환함수
    Dim userPathAndFile As String
    Dim wb As Workbook
    On Error Resume Next

    Set wb = ActiveWorkbook
    If wb.path <> "" Then
        ' 로컬 또는 네트워크 드라이브에서 저장된 경우
        userPathAndFile = wb.FullName
    Else
        ' SharePoint에 저장된 파일인 경우
        userPathAndFile = wb.FullName
        ' URL에서 "http://" 또는 "https://" 제거
        If Left(userPathAndFile, 7) = "http://" Then
            userPathAndFile = Mid(userPathAndFile, 8)
        ElseIf Left(userPathAndFile, 8) = "https://" Then
            userPathAndFile = Mid(userPathAndFile, 9)
        End If
    End If
    GetUserPath = userPathAndFile
End Function

Public Function GetUserMail() As String ' 이메일 반환함수
    On Error Resume Next
    Select Case UserGmail
    Case Is <> vbNullString: GetUserMail = UserGmail
    Case vbNullString: GetUserMail = UserOutlookMail
    End Select
End Function

Public Function UserGmail() As String
    On Error Resume Next
    UserGmail = CreateObject("WScript.Shell").RegRead("HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Common\Identity\ADUserName")
End Function

Public Function UserOutlookMail() As String
    On Error Resume Next
    Dim outApp As Object, outSession As Object
    Set outApp = CreateObject("Outlook.Application")
    If outApp Is Nothing Then UserOutlookMail = vbNullString: Exit Function
    Set outSession = outApp.Session.CurrentUser
    UserOutlookMail = outSession.AddressEntry.GetExchangeUser().PrimarySmtpAddress
    Set outApp = Nothing
End Function

' ==================== PERMISSION VALIDATION ====================
Public Function IsPermittedEmail() As Boolean ' 이메일로 사용자 권한여부 파악
   On Error Resume Next
   Dim permittedDomains As Variant
   Dim domain As Variant
   ' HRE: Added @hre.com domain for HRE users
   permittedDomains = Array("@pwc.com", "@bepsolar.com", "@hre.com")
   For Each domain In permittedDomains
       If InStr(GetUserMail(), domain) > 0 Then
           IsPermittedEmail = True
           Exit Function
       End If
   Next
   IsPermittedEmail = False
End Function

Public Function IsExpired() As Boolean
    IsExpired = (Format(Now(), "yyyy-mm-dd") >= ExpDate())
End Function

Public Sub ValidatePermission()
    On Error Resume Next
    If Not IsPermittedEmail() Then
        ' Currently disabled - uncomment to enforce email validation
        ' GoEnd "허가된 사용자 계정이 아닙니다!"
    End If
    If IsExpired() Then
        GoEnd "사용 기간이 만료되었습니다!"
    End If
End Sub

' ==================== FORM POSITIONING ====================
Sub FramePosition(ByVal Frame As Object)
    Dim sngLeft As Single, sngTop As Single
    With Frame
        .StartUpPosition = 0
        .Caption = AppName & " " & AppType
        Call ReturnPosition_CenterScreen(.Height, .Width, sngLeft, sngTop)
        .Left = sngLeft: .Top = sngTop
    End With
End Sub

Public Sub ReturnPosition_CenterScreen(ByVal sngHeight As Single, _
                                       ByVal sngWidth As Single, _
                                       ByRef sngLeft As Single, _
                                       ByRef sngTop As Single)
    Dim sngAppWidth As Single
    Dim sngAppHeight As Single
    Dim hwnd As Long
    Dim lreturn As Long
    Dim lpRect As udtRECT

    hwnd = Application.hwnd
    lreturn = GetWindowRect(hwnd, lpRect)
    sngAppWidth = ConvertPixelsToPoints(lpRect.Right - lpRect.Left, "X")
    sngAppHeight = ConvertPixelsToPoints(lpRect.Bottom - lpRect.Top, "Y")
    sngLeft = ConvertPixelsToPoints(lpRect.Left, "X") + ((sngAppWidth - sngWidth) / 2)
    sngTop = ConvertPixelsToPoints(lpRect.Top, "Y") + ((sngAppHeight - sngHeight) / 2)

End Sub

Public Function ConvertPixelsToPoints(ByVal sngPixels As Single, _
                                      ByVal sXorY As String) As Single
    Dim hDC As Long
    hDC = GetDC(0)
    If sXorY = "X" Then
       ConvertPixelsToPoints = sngPixels * (72 / GetDeviceCaps(hDC, 88))
    End If
    If sXorY = "Y" Then
       ConvertPixelsToPoints = sngPixels * (72 / GetDeviceCaps(hDC, 90))
    End If
    Call ReleaseDC(0, hDC)
End Function

' ==================== PROGRESS BAR ====================
Sub OpenProgress(ByVal StrMsg As String, Optional NotZero As Boolean)
    On Error Resume Next
    Dim FrmColor As Long
    If FrmProgress.Visible = True Then Exit Sub
    With FrmProgress
        Select Case Rnd * 100
            Case Is < 20: FrmColor = PwCYlw
            Case Is < 40: FrmColor = PwCTang
            Case Is < 60: FrmColor = PwCOrg
            Case Is < 80: FrmColor = PwCRose
            Case Else: FrmColor = PwCRed
        End Select
        .Show 0
        .LblMsg = StrMsg: .ProgressBar.BackColor = FrmColor
    End With
    Call FramePosition(FrmProgress)
    If NotZero = False Then Call CalculateProgress(0)
End Sub

Sub CalculateProgress(ByVal Progress As Single, Optional StrMsg As String)
    On Error Resume Next
    If Progress >= 1 Then Unload FrmProgress: Exit Sub
    If FrmProgress.Visible = False Then Call OpenProgress(StrMsg)
    With FrmProgress
        If StrMsg <> vbNullString Then .LblMsg.Caption = StrMsg
        .FraRate.Caption = Format(Progress, "0.0%")
        .ProgressBar.Width = Progress * .FraBar.Width
        .Repaint
    End With
End Sub

' ==================== PWC BRAND COLORS ====================
Public Function PwCGreenText() As Long
    PwCGreenText = RGB(0, 97, 0)
End Function

Public Function LightGreen() As Long
    LightGreen = RGB(198, 239, 206)
End Function

Public Function PwCLightGreen() As Long
    PwCLightGreen = RGB(196, 252, 159)
End Function

Public Function PwCLightYellow() As Long
    PwCLightYellow = RGB(255, 236, 189)
End Function

Public Function PwCLightRed() As Long
    PwCLightRed = RGB(247, 200, 196)
End Function

Public Function PwCLightOrg() As Long
    PwCLightOrg = RGB(254, 183, 145)
End Function

Public Function PwCLightTang() As Long
    PwCLightTang = RGB(255, 220, 169)
End Function

Public Function PwCLightBlue() As Long
    PwCLightBlue = RGB(179, 220, 249)
End Function

Public Function XBRLBlue() As Long
    XBRLBlue = RGB(153, 204, 255)
End Function

Public Function DarkBlue() As Long
    DarkBlue = RGB(0, 61, 171)
End Function

Public Function DarkGrey() As Long
    DarkGrey = RGB(45, 45, 45)
End Function

Public Function MediumGrey() As Long
    MediumGrey = RGB(70, 70, 70)
End Function

Public Function LightGrey() As Long
    LightGrey = RGB(222, 222, 222)
End Function

Public Function PwCRed() As Long
    PwCRed = RGB(224, 48, 30)
End Function

Public Function PwCOrg() As Long
    PwCOrg = RGB(208, 74, 2)
End Function

Public Function PwCTang() As Long
    PwCTang = RGB(235, 140, 0)
End Function

Public Function PwCYlw() As Long
    PwCYlw = RGB(255, 182, 0)
End Function

Public Function PwCRose() As Long
    PwCRose = RGB(219, 83, 106)
End Function

Public Function PwCLightRose() As Long
    PwCLightRose = RGB(241, 186, 195)
End Function

Public Function PwCDigitalRose() As Long
    PwCDigitalRose = RGB(217, 57, 84)
End Function

Public Function PwCGreen() As Long
    PwCGreen = RGB(23, 92, 44)
End Function

Public Function PwCBlue() As Long
    PwCBlue = RGB(0, 137, 235)
End Function

Public Function PwCDarkBlue() As Long
    PwCDarkBlue = RGB(0, 61, 171)
End Function
