REM  *****  BASIC  *****
'10分ごとに計測される環境データから、1日ごとの平均気温・CO2・湿度などを算出'
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
	
	Dim maxTemp As Double
	maxTemp = 0
	
	Dim minTemp As Double
	minTemp = 0
	
	Dim flag As Long
	flag = 0

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
	
	
	'*************************************************************************************************'
	'プログラムA'
	'1日ごとの環境の平均（CO2や湿度、気温）を算出するプログラム'
	For Each c In Range("A2:A16623")
		If InStr(c.Value, Cells(dateRow, dateCol).Value) > 0 Then
			ave = ave + Cells(row, 2).Value 'CO2に着目なう<------可変'
			count = count + 1
			
			'日付が変わる時'
			If Instr(Cells(row + 1, 1).Value, Cells(dateRow + 1, dateCol).Value) > 0 Then
				If Instr(Cells(row + 1, 1).Value, "0:00") > 0 Or Instr(Cells(row + 1, 1).Value, "0:10" ) > 0 Or Instr(Cells(row + 1, 1).Value, "0:20" ) > 0 Then
					Cells(dateRow, 8).Value = ave / count 'CO2に着目なう<-----可変'
					count = 0
					ave = 0
					dateRow = dateRow + 1
				End If
			End If
			
			row = row + 1
		End If
	Next c
	'*************************************************************************************************'

	
	'*************************************************************************************************'	
	'プログラムB'
	'1日ごとの最高・最低・差分気温を算出するプログラム ※先に平均気温とか求めてから！！！'
	For Each c In Range("A2:A16623")
		If InStr(c.Value, Cells(dateRow, dateCol).Value) > 0 Then
			If flag = 1 Then
				If Cells(row, 5).Value < minTemp Then
					minTemp = Cells(row, 5).Value
				End If
				
				If Cells(row, 5).Value > maxTemp Then
					maxTemp = Cells(row, 5).Value
				End If
				
				'日付が変わる時'
				If Instr(Cells(row + 1, 1).Value, Cells(dateRow + 1, dateCol).Value) > 0 Then
					If Instr(Cells(row + 1, 1).Value, "0:00") > 0 Or Instr(Cells(row + 1, 1).Value, "0:10" ) > 0 Or Instr(Cells(row + 1, 1).Value, "0:20" ) > 0 Then
						Cells(dateRow, 12).Value = maxTemp
						Cells(dateRow, 13).Value = minTemp
						Cells(dateRow, 14).Value = maxTemp - minTemp
						maxTemp = 0
						minTemp = 0
						flag = 2
						dateRow = dateRow + 1
					End If
				End If
			End If
			
			'0:00に計測されたときに、動く'
			If flag = 0 Then
				minTemp = Cells(row, 5).Value
				maxTemp = minTemp
				flag = 1
			End If
			

			If flag = 2 Then
				flag = 0
			End If

			row = row + 1
		End If
	Next c	
	'*************************************************************************************************'
	
End Sub
