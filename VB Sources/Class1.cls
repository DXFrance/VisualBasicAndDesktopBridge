VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private m_lShadowSize As Long
Private m_x As Long
Private m_y As Long
Private m_lWidth As Long
Private m_lHeight As Long
Private m_hWnd As Long
Private m_hWndOwner As Long

Private m_tBI As BITMAPINFO
Private m_lPtr As Long
Private m_hDC As Long
Private m_hDib As Long
Private m_hBmpOld As Long
Private m_lShadowColor As Long

Implements ISubclass


Private Property Let ISubclass_MsgResponse(ByVal RHS As EMsgResponse)
   '
End Property

Private Property Get ISubclass_MsgResponse() As EMsgResponse
   ISubclass_MsgResponse = emrPostProcess
End Property

Private Function ISubclass_WindowProc(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
   If (hwnd = m_hWnd) And (iMsg = WM_PAINT) Then
      Dim lHDC As Long
      Dim tR As RECT
      lHDC = GetDC(m_hWnd)
'      GetClientRect m_hWnd, tR
      BitBlt lHDC, tR.left, tR.top, tR.right - tR.left, tR.bottom - tR.top, m_hDC, 0, 0, vbSrcCopy
      ReleaseDC m_hWnd, lHDC
   End If
End Function
