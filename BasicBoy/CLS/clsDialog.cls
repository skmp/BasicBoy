VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private cdlg As OPENFILENAME
Private LastFileName As String

Public Function ShowOpen(ByVal Form_hWnd As Long, ByVal Title As String, ByVal InitDir As String, ByVal Filter As String) As String
  Dim i As Integer
  If Filter = "" Then Filter = "All Files (*.*)"
  Filter = Replace(Filter, "|", Chr(0))
  If Right(Filter, 1) <> Chr(0) Then Filter = Filter & Chr(0)
  If InitDir = "" Then InitDir = LastFileName & Chr(0)
  cdlg.lStructSize = Len(cdlg)
  cdlg.hwndOwner = Form_hWnd
  cdlg.hInstance = App.hInstance
  cdlg.lpstrFilter = Filter
  cdlg.lpstrFile = Space(254)
  cdlg.nMaxFile = 255
  cdlg.lpstrFileTitle = Space(254)
  cdlg.nMaxFileTitle = 255
  cdlg.lpstrInitialDir = InitDir & Chr(0)
  cdlg.lpstrTitle = Title
  cdlg.Flags = OFN_LONGNAMES Or OFN_FILEMUSTEXIST Or OFN_PATHMUSTEXIST
  ShowOpen = IIf(GetOpenFileName(cdlg), Trim(cdlg.lpstrFile), "")
  If Len(ShowOpen) > 0 Then LastFileName = ShowOpen
End Function

Public Function ShowSave(ByVal Form_hWnd As Long, ByVal Title As String, ByVal InitDir As String, ByVal Filter As String, ByVal DefExt As String) As String
  If Filter = "" Then Filter = "All Files (*.*)"
  Dim i As Integer
  For i = 1 To Len(Filter)
    If Mid(Filter, i, 1) = "|" Then Mid(Filter, i, 1) = Chr(0)
  Next i
  If Mid(Filter, Len(Filter), 1) <> Chr(0) Then Filter = Filter & Chr(0)
  If InitDir = "" Then InitDir = LastFileName
  cdlg.lStructSize = Len(cdlg)
  cdlg.lpstrTitle = Title
  cdlg.hwndOwner = Form_hWnd
  cdlg.hInstance = App.hInstance
  cdlg.lpstrFilter = Filter
  cdlg.lpstrDefExt = DefExt
  cdlg.lpstrFile = Space(254)
  cdlg.nMaxFile = 255
  cdlg.lpstrFileTitle = Space(254)
  cdlg.nMaxFileTitle = 255
  cdlg.lpstrInitialDir = InitDir
  cdlg.Flags = OFN_OVERWRITEPROMPT Or OFN_LONGNAMES Or OFN_NODEREFERENCELINKS Or OFN_PATHMUSTEXIST
  ShowSave = IIf(GetSaveFileName(cdlg), Trim(cdlg.lpstrFile), "")
  If Len(ShowSave) > 0 Then LastFileName = ShowSave
End Function

Public Property Let FileName(szFileName)
  LastFileName = szFileName
End Property

Public Property Get FileName()
  FileName = LastFileName
End Property

