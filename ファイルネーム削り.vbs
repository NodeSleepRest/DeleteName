Option Explicit

Dim fso ' Scripting.FileSystemObject
Dim f ' Scripting.File
Dim args
Dim strFileName
Dim intDelCount

With WScript
	Set args = .Arguments
	If args.Count < 1 Then .Quit
	intDelCount = InputBox("�t�@�C�����̍�����폜���镶��������͂��Ă��������B")
	If Len(intDelCount) = 0 OR Not IsNumeric(intDelCount) Then .Quit
End With

Set fso = CreateObject("Scripting.FileSystemObject")

On Error Resume Next

For Each strFileName in args

	set f = fso.GetFile(strFileName)
	f.Name = Mid(f.Name, intDelCount + 1)

	With Err
		Select Case .Number
		Case 58
			MsgBox "�������O�̃t�@�C�������݂��邽�ߏ����𒆒f���܂��B"
			Exit For
		Case 0
		' �G���[���������Ȃ������ꍇ�͉������Ȃ�
		Case Else
			MsgBox .Description & .Number
			.Clear
		End Select
	End With

Next

Set f = Nothing
Set fso = Nothing
