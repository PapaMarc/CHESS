Attribute VB_Name = "TaskListProcs"
Option Explicit

Declare Function GetWindow% Lib "USER" (ByVal hWnd%, ByVal wCmd%)
Declare Function ShowWindow Lib "USER" (ByVal hWnd As Integer, ByVal nCmdShow As Integer) As Integer

Global Const kAddTask = True
Global Const kRemoveTask = False
Private fAlwaysInTaskList As Boolean
Private fTaskHidden As Boolean

Global Const SW_HIDE = 0
Global Const SW_SHOW = 5
Global Const GW_OWNER = 4
Public Sub UpdateTasklist(fAdd As Boolean, frmCurr As Form)
    ' adds or removes application from windows tasklist
    Dim frm
    Dim fLoaded As Boolean
    
    Exit Sub
    If Not fAlwaysInTaskList Then
        If fAdd Then
            ' add to tasklist
            Call HideOwner(False)
        Else
            ' remove from tasklist
            If Not fTaskHidden Then

                For Each frm In Forms
                    If frm.hWnd <> frmCurr.hWnd And frm.hWnd <> frmAdmin.hWnd Then
                        fLoaded = True
                        Exit For
                    End If
                Next frm

                'If Not fLoaded Then
                    Call HideOwner(True)
                'End If
            End If
        End If
    End If
End Sub
Public Sub HideOwner(fHide As Boolean)
    ' hides or shows owner window of app
    Static hWndOwner As Integer
    Dim ret As Integer
                    
    If fTaskHidden <> fHide Then
        If hWndOwner = 0 Then
            hWndOwner = GetWindow(frmAdmin.hWnd, GW_OWNER)
        End If
        
        ret = ShowWindow(hWndOwner, IIf(fHide, SW_HIDE, SW_SHOW))
        fTaskHidden = fHide
    End If
End Sub
