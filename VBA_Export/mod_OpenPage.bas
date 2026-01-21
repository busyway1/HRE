Attribute VB_Name = "mod_OpenPage"
' ============================================================================
' Module: mod_OpenPage
' Project: HRE 연결마스터 (Consolidation Master)
' Migrated from: BEP v1.98
' Migration Date: 2026-01-21
'
' Description: Open external pages (SPO, Google Forms, Manual)
' Changes: Updated URLs for HRE version
' HRE Note: Manual URL and Google Form URL will need to be updated
' ============================================================================
Option Explicit
Sub OpenSPO()
    On Error Resume Next
    Dim URL As String
    URL = HideSheet.Range("E2").Value

    If URL = "" Then
        Msg "SPO 홈페이지 값을 설정해주세요!", vbExclamation
    End If
    Shell "cmd /c start " & URL, vbHide
End Sub
Sub OpenGoogleForm()
    On Error Resume Next
    Dim URL As String
    ' HRE: Update this URL with actual HRE Google Form when available
    URL = "https://docs.google.com/forms/d/e/1FAIpQLScJnpTjS1_1IPe3K4VE3hUlhRr8X0zpLt7uS7hPvaXBR4qQiA/viewform?usp=sf_link"
    Shell "cmd /c start " & URL, vbHide
End Sub
Sub OpenManual()
    On Error Resume Next
    Dim URL As String
    ' HRE: Update this URL with actual HRE manual URL when available
    URL = "https://www.notion.so/HRE-Manual-1503b6c9c09580c3a943fa26ab539e48?pvs=4"
    Shell "cmd /c start " & URL, vbHide
End Sub
