Attribute VB_Name = "mod_z_Module_GetCursor"
' ============================================================================
' Module: mod_z_Module_GetCursor
' Project: HRE 연결마스터 (Consolidation Master)
' Migrated from: BEP v1.98
' Migration Date: 2026-01-21
'
' Description: Cursor position and DPI handling utilities
' Changes: COPY AS-IS - No HRE-specific adaptations required
' ============================================================================
Option Explicit

#If VBA7 Then
Declare PtrSafe Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
#Else
Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetDeviceCaps Lib "Gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
#End If

Type POINTAPI
    X As Long
    Y As Long
End Type

Const LOGPIXELSX = 88
Const LOGPIXELSY = 90

Public Function pointsPerPixelX() As Double
    Dim hDC As Long
    hDC = GetDC(0)
    pointsPerPixelX = 72 / GetDeviceCaps(hDC, LOGPIXELSX)
    ReleaseDC 0, hDC
End Function

Public Function pointsPerPixelY() As Double
    Dim hDC As Long
    hDC = GetDC(0)
    pointsPerPixelY = 72 / GetDeviceCaps(hDC, LOGPIXELSY)
    ReleaseDC 0, hDC
End Function

Public Function WhereIsTheMouseAt() As POINTAPI
    Dim mPos As POINTAPI
    GetCursorPos mPos
    WhereIsTheMouseAt = mPos
End Function

Public Function convertMouseToForm() As POINTAPI
    Dim mPos As POINTAPI
    mPos = WhereIsTheMouseAt
    mPos.X = pointsPerPixelY * mPos.X
    mPos.Y = pointsPerPixelX * mPos.Y
    convertMouseToForm = mPos
End Function
