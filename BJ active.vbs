'--------------------------------------------------------
'SetMpcTimeStretchFilter.vbs
'--------------------------------------------------------
Option Explicit
Const PROC_NAME = "BlueJeans.exe"
Const KEYSEND_SLEEP = 100
				

Dim m_objWshShell
Set m_objWshShell = CreateObject("Wscript.Shell")


Dim objWshShell
'シェルオブジェクトの作成
Set objWshShell = WScript.CreateObject("WScript.Shell")
'シェルの実行
objWshShell.Run """C:\Users\testtets"""

Call Wscript.Sleep(8000)

'メイン処理呼出
MainProc()

'メイン処理
Sub MainProc()
	'プロセスID取得
	Dim procId
	Dim procIdParentArray
	procIdParentArray = GetProcId(PROC_NAME)
	
	Dim element
	Dim temp
	for each element in procIdParentArray
		temp = temp & element & vbCrLf
	next
	
	procId = parentIdManyMostPopNumber(procIdParentArray)
	'アクティブ化
	Call m_objWshShell.AppActivate(procId)
	
	' 'キー送信前に少し待機
	Call Wscript.Sleep(100)
	' 'キー送信
	Call SendKeys("{tab}")
	Call SendKeys("{tab}")
	Call SendKeys("{tab}")	
	Call SendKeys("{tab}")
	Call SendKeys("{ENTER}")
End Sub

'親のプロセスIDを取得
Function GetProcId(procName)
	Dim Service
	Dim QfeSet
	Dim Qfe
	Dim ParentProcId()
	ReDim ParentProcId(0)
	Dim ParentProcessIdCounter
	ParentProcessIdCounter = 0
	Set Service = WScript.CreateObject("WbemScripting.SWbemLocator").ConnectServer
	Set QfeSet = Service.ExecQuery("SELECT * FROM Win32_Process WHERE Caption = '"& procName &"'")
	Dim temp
	For Each Qfe in QfeSet
		temp = temp & Qfe.ProcessId & " : " &Qfe.ParentProcessId & vbCrLf
		
		ReDim Preserve ParentProcId(ParentProcessIdCounter)
		ParentProcId(ParentProcessIdCounter) = Qfe.ParentProcessId
		ParentProcessIdCounter = ParentProcessIdCounter + 1
	Next
	' msgbox Join(ParentProcId,",")
	' msgbox temp
	GetProcId = ParentProcId
End Function

'一番出現数の多い親のプロセスIDを取得
Function parentIdManyMostPopNumber(parentIdArray)
	Dim arrayLength
	arrayLength = UBound(parentIdArray)
	Dim myDictionary
	Set myDictionary = CreateObject("Scripting.Dictionary")

	Dim parentId
	Dim returnId
	For Each parentId in parentIdArray
		' あるキーが存在するかどうかを判定
		If myDictionary.Exists(parentId) Then
			' 存在する場合 値を1増加
			myDictionary(parentId) = myDictionary(parentId) + 1
		else
			myDictionary.Add parentId,1
		End If
		returnId = parentId
	Next
	
	Dim str
	Dim oneDictonary
	Dim maxPopNumber
	Dim maxPopCountKey
	maxPopNumber = 0
	For Each oneDictonary In myDictionary
		if maxPopNumber <= myDictionary(oneDictonary) Then
			maxPopNumber = myDictionary(oneDictonary) ' 出現数が今までより大きいならそのキーを一時保存
			maxPopCountKey = oneDictonary ' 出現数が今までの中で一番多いもののプロセスIDを一時保存
		end if
		str = str & oneDictonary & " : " & myDictionary(oneDictonary) & vbCrLf
	Next
	
	parentIdManyMostPopNumber = returnId
End Function

'キー送信
Sub SendKeys(key)
	Call Wscript.Sleep(KEYSEND_SLEEP)
	Call m_objWshShell.SendKeys(key)
End Sub