'--------------------------------------------------------
'SetMpcTimeStretchFilter.vbs
'--------------------------------------------------------
Option Explicit

Const SET_PLAYBACK_RATE = "1.5"
Const PROC_NAME = "mpc-hc.exe"
Const KEYSEND_SLEEP = 20

Dim m_objWshShell
Set m_objWshShell = CreateObject("Wscript.Shell")

'メイン処理呼出
MainProc()

'メイン処理
Sub MainProc()
    'プロセスID取得
    Dim procId
    procID = GetProcId(PROC_NAME)
    'プロセスID未取得で終了
    If procID = 0 Then
        Exit Sub
    End If
    'アクティブ化
    Call m_objWshShell.AppActivate(procId)
    'キー送信前に少し待機
    Call Wscript.Sleep(100)
    'キー送信
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

'プロセスID取得
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

'キー送信
Sub SendKeys(key)
    Call Wscript.Sleep(KEYSEND_SLEEP)
    Call m_objWshShell.SendKeys(key)
End Sub