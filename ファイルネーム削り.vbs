Option Explicit

Dim fso ' Scripting.FileSystemObject
Dim f ' Scripting.File
Dim args
Dim strFileName
Dim intDelCount

With WScript
	Set args = .Arguments
	If args.Count < 1 Then .Quit
	intDelCount = InputBox("ファイル名の左から削除する文字数を入力してください。")
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
			MsgBox "同じ名前のファイルが存在するため処理を中断します。"
			Exit For
		Case 0
		' エラーが発生しなかった場合は何もしない
		Case Else
			MsgBox .Description & .Number
			.Clear
		End Select
	End With

Next

Set f = Nothing
Set fso = Nothing
