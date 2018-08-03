'--------------------------------------------------------
'SetMpcTimeStretchFilter.vbs
'--------------------------------------------------------
Option Explicit

Const SET_PLAYBACK_RATE = "1.5"
Const PROC_NAME = "mpc-hc.exe"
Const KEYSEND_SLEEP = 20

Dim m_objWshShell
Set m_objWshShell = CreateObject("Wscript.Shell")

'���C�������ďo
MainProc()

'���C������
Sub MainProc()
    '�v���Z�XID�擾
    Dim procId
    procID = GetProcId(PROC_NAME)
    '�v���Z�XID���擾�ŏI��
    If procID = 0 Then
        Exit Sub
    End If
    '�A�N�e�B�u��
    Call m_objWshShell.AppActivate(procId)
    '�L�[���M�O�ɏ����ҋ@
    Call Wscript.Sleep(100)
    '�L�[���M
    Call SendKeys("%")
    Call SendKeys("p")
    Call SendKeys("f")
    'Call SendKeys("c")
    Call SendKeys("{DOWN}")
    Call SendKeys("{DOWN}")
    Call SendKeys("{DOWN}")
    Call SendKeys("{DOWN}")
    Call SendKeys("{DOWN}")
    Call SendKeys("{ENTER}")
    Call SendKeys(SET_PLAYBACK_RATE)
    Call SendKeys("{ENTER}")
End Sub

'�v���Z�XID�擾
Function GetProcId(procName)
    Dim Service
    Dim QfeSet
    Dim Qfe
    Set Service = WScript.CreateObject("WbemScripting.SWbemLocator").ConnectServer
    Set QfeSet = Service.ExecQuery("SELECT * FROM Win32_Process WHERE Caption = '"& procName &"'")
    GetProcId = 0
    For Each Qfe in QfeSet
        GetProcId = Qfe.ProcessId
        Exit For
    Next
End Function

'�L�[���M
Sub SendKeys(key)
    Call Wscript.Sleep(KEYSEND_SLEEP)
    Call m_objWshShell.SendKeys(key)
End Sub