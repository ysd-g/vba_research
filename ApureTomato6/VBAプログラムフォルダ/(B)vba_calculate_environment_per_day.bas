REM  *****  BASIC  *****
'10分ごとに計測される環境データから、1日ごとの平均気温・CO2・湿度などを算出'
'基本的には、7列目に日にち（date）を新しく設け、★の部分の数字さえ変えればよい'
Option VBASupport 1

Sub Main

	Dim row As Long '行'
	row = 2 '2行目から始まる'
	
	Dim dateRow As Long '日にち行'
	dateRow = 2
	
	Dim dateCol As Long '列'
	dateCol = 7 '2列目'

	Dim ave As Double '平均値'
	ave = 0
	
	Dim count As Long
	count = 0
	
	'-----------------------------------------------------------'
	'0:00 とか 0:10が何個含まれているか'
	'For Each c In Range("A2:A16623")
	'	If InStr(c.Value, "0:20") > 0 Then
	'		count = count + 1
	'	End If
	'Next c
	'Cells(4, 16).Value = count
	'count = 0
	'-----------------------------------------------------------'
	
	'1日ごとのCO2や湿度、気温の平均を算出するプログラム'
	For Each c In Range("A2:A16623")
		If InStr(c.Value, Cells(dateRow, dateCol).Value) > 0 Then
			ave = ave + Cells(row, 2).Value 'CO2に着目なう<------可変★'
			count = count + 1
			
			'日付が変わる時'
			If Instr(Cells(row + 1, 1).Value, Cells(dateRow + 1, dateCol).Value) > 0 Then
				If Instr(Cells(row + 1, 1).Value, "0:00") > 0 Or Instr(Cells(row + 1, 1).Value, "0:10" ) > 0 Or Instr(Cells(row + 1, 1).Value, "0:20" ) > 0 Then
					Cells(dateRow, 8).Value = ave / count 'CO2に着目なう<-----可変★'
					count = 0
					ave = 0
					dateRow = dateRow + 1
				End If
			End If
			
			row = row + 1
		End If
	Next c

End Sub
