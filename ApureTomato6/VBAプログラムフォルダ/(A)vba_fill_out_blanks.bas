REM  *****  BASIC  *****
'気温や湿度などの生の環境データで、空欄になっている箇所を平均で埋めるプログラム'
Option VBASupport 1


'このプログラムを動かす前に、スプレッドシートにデータを移して、
'空欄のところを0に置き換えて、またodsファイルに移してVBA実行しないとエラーになる'

Sub Main
	Dim row As Long '行'
	row = 2 '2行目から始まる'

	Dim col As Long '列'
	col = 2 '2列目'

	Dim ave As Double '平均値'
	ave = 0

	For row = 2 To 16623
		If Cells(row, col).Value = 0 Then '気温の列を選択なう'
			ave = (Cells(row - 1, col).Value + Cells(row + 1, col).Value) / 2
			Cells(row, col).Value = ave
		End If
	Next row
End Sub
