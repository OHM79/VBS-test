Option Explicit

'WMI�ɂĎg�p����e��I�u�W�F�N�g���`�E��������B
Dim oClassSet
Dim oClass
Dim oLocator
Dim oService
Dim sMesStr

'���[�J���R���s���[�^�ɐڑ�����B
Set oLocator = WScript.CreateObject("WbemScripting.SWbemLocator")
Set oService = oLocator.ConnectServer
'�N�G���[������ WQL �ɂĎw�肷��B
Set oClassSet = oService.ExecQuery("Select * From Win32_Process")

' '�R���N�V��������͂���B
' For Each oClass In oClassSet

' sMesStr = sMesStr & oClass.Caption & vbCrLf

' Next
For Each oClass In oClassSet

sMesStr = sMesStr & oClass.Description & " : " & oClass.Caption & " : " & oClass.ParentProcessId & vbCrLf

Next

MsgBox "�T�[�r�X�Ɋւ�����ł��B" & vbCrLf & vbCrLf & sMesStr

'�g�p�����e��I�u�W�F�N�g����Еt������B
Set oClassSet = Nothing
Set oClass = Nothing
Set oService = Nothing
Set oLocator = Nothing