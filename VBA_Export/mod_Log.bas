Attribute VB_Name = "mod_Log"
' ============================================================================
' Module: mod_Log
' Project: HRE 연결마스터 (Consolidation Master)
' Migrated from: BEP v1.98
' Migration Date: 2026-01-21
'
' Description: Usage logging via Google Forms
' Changes: Updated Google Forms URLs for HRE version
' HRE Note: Forms URLs will need to be replaced with actual HRE forms
' ============================================================================
Option Explicit
Private Declare PtrSafe Function WideCharToMultiByte Lib "kernel32" ( _
    ByVal CodePage As Long, _
    ByVal dwFlags As Long, _
    ByVal lpWideCharStr As LongPtr, _
    ByVal cchWideChar As Long, _
    ByVal lpMultiByteStr As LongPtr, _
    ByVal cbMultiByte As Long, _
    ByVal lpDefaultChar As LongPtr, _
    ByVal lpUsedDefaultChar As LongPtr) As Long
Private Const CP_UTF8 As Long = 65001
Public Sub LogData(Answer1 As String, Answer2 As String)
    Dim formUrl As String
    Dim formdata As String
    Dim Answer3 As String
    Dim Answer4 As String
    On Error Resume Next
    Answer3 = GetUserInfo()
    Answer4 = GetUserPath()
    ' HRE: Update this URL with actual HRE Google Form when available
    formUrl = "https://docs.google.com/forms/d/e/1FAIpQLSdttj0I6RrDD4Y33rC_e4Co_zI_WKyUkw2HR18e2HfRB5JYWw/formResponse"
    formdata = "entry.1826767042=" & UTF8_URLEncode(Answer1) & "&" & _
              "entry.476426559=" & UTF8_URLEncode(Answer2) & "&" & _
              "entry.791850601=" & UTF8_URLEncode(Answer3) & "&" & _
              "entry.1798711265=" & UTF8_URLEncode(Answer4)
    Dim xhr As Object
    Set xhr = CreateObject("MSXML2.ServerXMLHTTP")
    xhr.Open "POST", formUrl, False
    xhr.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    xhr.send formdata
End Sub
Public Sub LogData_Access(Answer1 As String, Answer2 As String)
    Dim formUrl As String
    Dim formdata As String
    Dim Answer3 As String
    Dim Answer4 As String
    Dim Answer5 As String
    On Error Resume Next
    Answer3 = GetUserInfo()
    Answer4 = GetUserMail()
    Answer5 = GetUserPath()
    ' HRE: Update this URL with actual HRE Google Form when available
    formUrl = "https://docs.google.com/forms/d/e/1FAIpQLScvuIkA1dtTfpVulaSYMNUhkJdyj3q6A-iykWF4YePzQHs4Dg/formResponse"
    formdata = "entry.988015572=" & UTF8_URLEncode(Answer1) & "&" & _
              "entry.1962065207=" & UTF8_URLEncode(Answer2) & "&" & _
              "entry.171765383=" & UTF8_URLEncode(Answer3) & "&" & _
              "entry.2134702071=" & UTF8_URLEncode(Answer4) & "&" & _
              "entry.1996784553=" & UTF8_URLEncode(Answer5)
    Dim xhr As Object
    Set xhr = CreateObject("MSXML2.ServerXMLHTTP")
    xhr.Open "POST", formUrl, False
    xhr.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    xhr.send formdata
End Sub
Private Function UTF8_URLEncode(ByVal text As String) As String
    Dim utf8() As Byte
    Dim length As Long
    Dim i As Long
    Dim result As String
    On Error Resume Next
    ' Convert to UTF-8
    length = WideCharToMultiByte(CP_UTF8, 0, StrPtr(text), -1, 0, 0, 0, 0) - 1
    If length > 0 Then
        ReDim utf8(0 To length - 1)
        WideCharToMultiByte CP_UTF8, 0, StrPtr(text), -1, VarPtr(utf8(0)), length, 0, 0
    End If
    ' URL encode
    For i = 0 To length - 1
          Select Case utf8(i)
            Case 65 To 90, 97 To 122, 48 To 57, 45, 46, 95, 126  ' A-Z, a-z, 0-9, "-", ".", "_", "~"
                result = result & Chr(utf8(i))
            Case 32  ' Space
                result = result & "+"
            Case Else
                result = result & "%" & Right("0" & Hex(utf8(i)), 2)
        End Select
    Next i
    UTF8_URLEncode = result
End Function
