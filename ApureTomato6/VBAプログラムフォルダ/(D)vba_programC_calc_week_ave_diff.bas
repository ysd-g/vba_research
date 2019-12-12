REM  *****  BASIC  *****
Option VBASupport 1

Sub Main

	Dim row As Long '行'
	row = 2 '2行目から始まる'
	
	Dim dateRow As Long '日にち行'
	dateRow = 2
	
	Dim dateCol As Long '列'
	dateCol = 7 '2列目'
	
	Dim i As Integer
	i = 1

	Dim ave As Double '平均値'
	ave = 0
	
	Dim count As Long
	count = 0
	
	Dim AveCO27 As Double
	AveCO27 = 0
	
	Dim AveSatu7 As Double
	AveSatu7 = 0
	
	Dim AveHum7 As Double
	AveHum7 = 0
	
	Dim AveTemp7 As Double
	AveTemp7 = 0
	
	Dim AveMaxTemp7 As Double
	AveMaxTemp7 = 0
	
	Dim AveMinTemp7 As Double
	AveMinTemp7 = 0
	
	Dim AveDiffTemp7 As Double
	AveDiffTemp7 = 0
	
	Dim SumMaxTemp7 As Double
	SumMaxTemp7 = 0
	
	Dim SumMinTemp7 As Double
	SumMinTemp7 = 0
	
	Dim SumDiffTemp7 As Double
	SumDiffTemp7 = 0
	
	Dim currentDate As String
	currentDate = ""
	
	Dim house As Long
	house = 0
	
	Dim maxTemp As Double
	maxTemp = 0
	
	Dim minTemp As Double
	minTemp = 0
	
	Dim flag As Long
	flag = 0
	
	Dim AveCO27_col As Long
	AveCO27_col = 0
	
	Dim AveSatu7_col As Long
	AveSatu7_col = 0
	
	Dim AveHum7_col As Long
	AveHum7_col = 0
	
	Dim AveTemp7_col As Long
	AveTemp7_col = 0
	
	Dim MaxTemp7_col As Long
	MaxTemp7_col = 0
	
	Dim MinTemp7_col As Long
	MinTemp7_col = 0
	
	Dim DiffTemp7_col As Long
	DiffTemp7_col = 0
	

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
	'プログラムC'
	'正規化データ > Aiko_ver2 > Aiko_normalization_ver2.odsの環境の1週間データを算出するプログラム
	'For Each c In Range("B2:B1526")
	For Each c In Range("B223:B1526") '*************可変★
		If flag = 1 Then
			If c.value = currentDate And house = Cells(c.Row, 4).Value Then
				Cells(c.Row, 11).Value = AveCO27
				Cells(c.Row, 12).Value = AveSatu7
				Cells(c.Row, 13).Value = AveHum7
				Cells(c.Row, 14).Value = AveTemp7
				Cells(c.Row, 15).Value = AveMaxTemp7
				Cells(c.Row, 16).Value = AveMinTemp7
				Cells(c.Row, 17).Value = AveDiffTemp7
				Cells(c.Row, 18).Value = SumMaxTemp7
				Cells(c.Row, 19).Value = SumMinTemp7
				Cells(c.Row, 20).Value = SumDiffTemp7
			Else
				AveCO27 = 0
				AveSatu7 = 0
				AveHum7 = 0
				AveTemp7 = 0
				AveMaxTemp7 = 0
				AveMinTemp7 = 0
				AveDiffTemp7 = 0
				SumMaxTemp7 = 0
				SumMinTemp7 = 0
				SumDiffTemp7 = 0
				flag = 0
			End If
		End If
		
		If flag = 0 Then 								'日付が変わった最初の処理
			currentDate = c.value					'2列目の該当する日付を代入
			house = Cells(c.Row, 4).Value	'ハウスの番号を得る
			
			If house = 1 Then						'ハウスNoが1なら30列目を対象とする
				dateCol = 30
			ElseIf house = 2 Then					'ハウスNoが2なら46列目を対象とする
				dateCol = 46
			End If
			
			For dateRow = 2 To 116				'*************可変★
																'30列目(1号棟)，もしくは46列目(2号棟)から該当する日付を探索する'
				If currentDate = Cells(dateRow, dateCol).Value Then
					AveCO27_col = dateCol + 1		'31列目(1号棟)，もしくは47列目(2号棟)
					AveSatu7_col = dateCol + 2		'32列目(1号棟)，もしくは48列目(2号棟)
					AveHum7_col = dateCol + 3		'33列目(1号棟)，もしくは49列目(2号棟)
					AveTemp7_col = dateCol + 4		'34列目(1号棟)，もしくは50列目(2号棟)
					MaxTemp7_col = dateCol + 5	'35列目(1号棟)，もしくは51列目(2号棟)
					MinTemp7_col = dateCol + 6		'36列目(1号棟)，もしくは52列目(2号棟)
					DiffTemp7_col = dateCol + 7		'37列目(1号棟)，もしくは53列目(2号棟)
				
					For i = 1 To 7
							AveCO27 = AveCO27 + Cells(dateRow - i, AveCO27_col).Value
							AveSatu7 = AveSatu7 + Cells(dateRow - i, AveSatu7_col).Value
							AveHum7 = AveHum7 + Cells(dateRow - i, AveHum7_col).Value
							AveTemp7 = AveTemp7 + Cells(dateRow - i, AveTemp7_col).Value
							AveMaxTemp7 = AveMaxTemp7 + Cells(dateRow - i, MaxTemp7_col).Value
							AveMinTemp7 = AveMinTemp7 + Cells(dateRow - i, MinTemp7_col).Value
							AveDiffTemp7 = AveDiffTemp7 + Cells(dateRow - i, DiffTemp7_col).Value
							SumMaxTemp7 = AveMaxTemp7
							SumMinTemp7 = AveMinTemp7
							SumDiffTemp7 = AveDiffTemp7
							
							If i = 7 Then
								AveCO27 = AveCO27 / 7
								AveSatu7 = AveSatu7 / 7
								AveHum7 = AveHum7 / 7
								AveTemp7 = AveTemp7 / 7
								AveMaxTemp7 = AveMaxTemp7 / 7
								AveMinTemp7 = AveMinTemp7 / 7
								AveDiffTemp7 = AveDiffTemp7 / 7
							End If
					Next i
					 
					Cells(c.Row, 11).Value = AveCO27
					Cells(c.Row, 12).Value = AveSatu7
					Cells(c.Row, 13).Value = AveHum7
					Cells(c.Row, 14).Value = AveTemp7
					Cells(c.Row, 15).Value = AveMaxTemp7
					Cells(c.Row, 16).Value = AveMinTemp7
					Cells(c.Row, 17).Value = AveDiffTemp7
					Cells(c.Row, 18).Value = SumMaxTemp7
					Cells(c.Row, 19).Value = SumMinTemp7
					Cells(c.Row, 20).Value = SumDiffTemp7
					flag = 1
					Exit For
				End If
			Next dateRow
			
		End If
	
	Next c
	
End Sub

