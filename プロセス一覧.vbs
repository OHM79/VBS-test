Option Explicit

'WMIにて使用する各種オブジェクトを定義・生成する。
Dim oClassSet
Dim oClass
Dim oLocator
Dim oService
Dim sMesStr

'ローカルコンピュータに接続する。
Set oLocator = WScript.CreateObject("WbemScripting.SWbemLocator")
Set oService = oLocator.ConnectServer
'クエリー条件を WQL にて指定する。
Set oClassSet = oService.ExecQuery("Select * From Win32_Process")

' 'コレクションを解析する。
' For Each oClass In oClassSet

' sMesStr = sMesStr & oClass.Caption & vbCrLf

' Next
For Each oClass In oClassSet

sMesStr = sMesStr & oClass.Description & " : " & oClass.Caption & " : " & oClass.ParentProcessId & vbCrLf

Next

MsgBox "サービスに関する情報です。" & vbCrLf & vbCrLf & sMesStr

'使用した各種オブジェクトを後片付けする。
Set oClassSet = Nothing
Set oClass = Nothing
Set oService = Nothing
Set oLocator = Nothing