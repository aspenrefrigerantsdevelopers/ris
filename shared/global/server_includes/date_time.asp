<%
'---- Returns a date formated in "m/d/yy" format ----
Function ShortDate(inDate)
	If Not IsNull(inDate) And inDate <> "" Then
		If IsDate(inDate) Then
			ShortDate = Month(inDate) & "/" & Day(inDate) & "/" & Mid(Year(inDate),3)
		Else
			'ShortDate = Month(Date()) & "/" & Day(Date()) & "/" & Mid(Year(Date()),3)
			ShortDate = ""
		End If
	Else
		'ShortDate = Month(Date()) & "/" & Day(Date()) & "/" & Mid(Year(Date()),3)
		ShortDate = ""
	End If
End Function

'---- Returns a time formated in "h:mm AM" format ----
Function ShortTime(inTime)
	If Not IsNull(inTime) And IsDate(inTime) Then 'Removed 10/13/05 because excluding date at midnight (no time) And InStr(1, inTime, ":") > 0 Then
		Dim Hr, Min, AmPm
		If Hour(inTime) = 0 Then
			Hr = 12
			AmPm = "AM"
		ElseIf Hour(inTime) < 12 Then
			Hr = Hour(inTime)
			AmPm = "AM"
		ElseIf Hour(inTime) = 12 Then
			Hr = Hour(inTime)
			AmPm = "PM"
		Else
			Hr = Hour(inTime) - 12
			AmPm = "PM"
		End If
		If Minute(inTime) < 10 Then
			Min = "0" & Minute(inTime)
		Else
			Min = Minute(inTime)
		End If
		ShortTime = Hr & ":" & Min & " " & AmPm
	Else
		ShortTime = ""
	End If
End Function

Function ShortDateTime(inDateTime)
	ShortDateTime = ShortDate(inDateTime) & " " & ShortTime(inDateTime)
End Function

Function MediumDate(inDateTime)
	If Not IsNull(inDateTime) And inDateTime <> "" And IsDate(inDateTime) Then
		MediumDate = Trim(Left(FormatDateTime(inDateTime, 1), 3) & " " & ShortDate(inDateTime))
	Else
		MediumDate = ""
	End If
End Function

Function MediumDateTime(inDateTime)
	If Not IsNull(inDateTime) And inDateTime <> "" And IsDate(inDateTime) Then
		MediumDateTime = Trim(Left(FormatDateTime(inDateTime, 1), 3) & " " & ShortDate(inDateTime) & " " & ShortTime(inDateTime))
	Else
		MediumDateTime = ""
	End If
End Function

Function LongDate(inDateTime)
	If Not IsNull(inDateTime) And inDateTime <> "" And IsDate(inDateTime) Then
		LongDate = Trim(FormatDateTime(inDateTime, 1))
	Else
		LongDate = ""
	End If
End Function

Function LongDateTime(inDateTime)
	If Not IsNull(inDateTime) And inDateTime <> "" And IsDate(inDateTime) Then
		LongDateTime = Trim(FormatDateTime(inDateTime, 1) & " " & ShortTime(inDateTime))
	Else
		LongDateTime = ""
	End If
End Function

'---- Converts AS/400 dates to PC format and vise versa ----
Function ConvertDate(inDate,inFormat,outFormat)
	Dim outDate
	If Not IsNull(inDate) Then
		Select Case inFormat
			Case "PC":
				If IsDate(inDate) Then
					Select Case outFormat
						Case "PC":
							ConvertDate = CDate(inDate)
						Case "CCYYMMDD":
							outDate = CStr(Year(inDate))
							If CInt(CStr(Month(inDate))) < 10 Then
								outDate = outDate & "0" & CStr(Month(inDate))
							Else
								outDate = outDate & CStr(Month(inDate))
							End If
							If CInt(CStr(Day(inDate))) < 10 Then
								outDate = outDate & "0" & CStr(Day(inDate))
							Else
								outDate = outDate & CStr(Day(inDate))
							End If
							ConvertDate = CLng(outDate)
						Case "YYMMDD":
							outDate = Mid(CStr(Year(inDate)),3,2)
							If CInt(CStr(Month(inDate))) < 10 Then
								outDate = outDate & "0" & CStr(Month(inDate))
							Else
								outDate = outDate & CStr(Month(inDate))
							End If
							If CInt(CStr(Day(inDate))) < 10 Then
								outDate = outDate & "0" & CStr(Day(inDate))
							Else
								outDate = outDate & CStr(Day(inDate))
							End If
							ConvertDate = outDate
						Case "MMDDYY":
							If CInt(CStr(Month(inDate))) < 10 Then
								outDate = "0" & CStr(Month(inDate))
							Else
								outDate = CStr(Month(inDate))
							End If
							If CInt(CStr(Day(inDate))) < 10 Then
								outDate = outDate & "0" & CStr(Day(inDate))
							Else
								outDate = outDate & CStr(Day(inDate))
							End If
							outDate = outDate & Mid(CStr(Year(inDate)),3,2)
							ConvertDate = outDate
						Case "MMDDCCYY":
							If CInt(CStr(Month(inDate))) < 10 Then
								outDate = "0" & CStr(Month(inDate))
							Else
								outDate = CStr(Month(inDate))
							End If
							If CInt(CStr(Day(inDate))) < 10 Then
								outDate = outDate & "0" & CStr(Day(inDate))
							Else
								outDate = outDate & CStr(Day(inDate))
							End If
							outDate = outDate & CStr(Year(inDate))
							ConvertDate = outDate
						Case "M-D-YY":
							ConvertDate = CStr(Month(inDate)) & "-" & CStr(Day(inDate)) & "-" & Mid(CStr(Year(inDate)),3,2)
						Case "CCYY-MM-DD"
							outDate = CStr(Year(inDate))
							If CInt(CStr(Month(inDate))) < 10 Then
								outDate = outDate & "-0" & CStr(Month(inDate))
							Else
								outDate = outDate & "-" & CStr(Month(inDate))
							End If
							If CInt(CStr(Day(inDate))) < 10 Then
								outDate = outDate & "-0" & CStr(Day(inDate))
							Else
								outDate = outDate & "-" & CStr(Day(inDate))
							End If
							ConvertDate = outDate
						Case "MMYY":
							If CInt(CStr(Month(inDate))) < 10 Then
								outDate = outDate & "0" & CStr(Month(inDate))
							Else
								outDate = outDate & CStr(Month(inDate))
							End If
							outDate = outDate & Mid(CStr(Year(inDate)), 3, 2)
							ConvertDate = outDate
						Case Else
							ConvertDate = ""
					End Select
				Else
					ConvertDate = ""
				End If
			Case "CCYYMMDD":
				'If Not IsNull(inDate) Then
					'If IsNumeric(inDate) Then
					'inDate = CDbl(inDate)
						If IsDate(Mid(CStr(inDate),5,2) & "-" & Mid(CStr(inDate),7,2) & "-" & Mid(CStr(inDate),1,4)) Then
							Select Case outFormat
								Case "PC":
									outDate = Mid(CStr(inDate),5,2) & "-" & Mid(CStr(inDate),7,2) & "-" & Mid(CStr(inDate),1,4)
									ConvertDate = CDate(outDate)
								Case "CCYYMMDD":
									ConvertDate = CStr(inDate)
								Case "YYMMDD":
									outDate = Mid(CStr(inDate),3,6)
									ConvertDate = CStr(outDate)
								Case "MMDDYY":
									outDate = Mid(CStr(inDate),5,2) & Mid(CStr(inDate),7,2) & Mid(CStr(inDate),3,2)
									ConvertDate = CStr(outDate)
								Case "MMYY":
									outDate = Mid(CStr(inDate),5,2) & Mid(CStr(inDate),3,2)
									ConvertDate = CStr(outDate)
								Case Else
									ConvertDate = ""
							End Select
						Else
							ConvertDate = ""
						End If
					'Else
					'	ConvertDate = ""
					'End If
				'Else
				'	ConvertDate = ""
				'End If
			Case "YYMMDD":
				If Len(CStr(inDate)) = 5 Then inDate = "0" & CStr(inDate)
				If IsDate(Mid(CStr(inDate),3,2) & "-" & Mid(CStr(inDate),5,2) & "-" & Mid(CStr(inDate),1,2)) Then
					Select Case outFormat
						Case "PC":
							outDate = Mid(CStr(inDate),3,2) & "-" & Mid(CStr(inDate),5,2) & "-" & Mid(CStr(inDate),1,2)
							ConvertDate = CDate(outDate)
						Case "CCYYMMDD":
							If Mid(Cstr(inDate),1,2) > 50 Then
								outDate = "19" & CStr(inDate)
							Else
								outDate = "20" & Cstr(inDate)
							End If
							ConvertDate = CStr(outDate)
						Case "YYMMDD":
							ConvertDate = CStr(inDate)
						Case "MMDDYY":
							outDate = Mid(CStr(inDate),3,2) & Mid(CStr(inDate),5,2) & Mid(CStr(inDate),1,2)
							ConvertDate = CStr(outDate)
						Case Else
							ConvertDate = ""
					End Select
				Else
					ConvertDate = ""
				End If
			Case "MMDDYY":
				If Len(CStr(inDate)) = 5 Then inDate = "0" & CStr(inDate)
				If IsDate(Mid(CStr(inDate),1,2) & "-" & Mid(CStr(inDate),3,2) & "-" & Mid(CStr(inDate),5,2)) Then
					Select Case outFormat
						Case "PC":
							outDate = Mid(CStr(inDate),1,2) & "-" & Mid(CStr(inDate),3,2) & "-" & Mid(CStr(inDate),5,2)
							ConvertDate = CDate(outDate)
						Case "CCYYMMDD":
							If Mid(Cstr(inDate),5,2) > 50 Then
								outDate = "19" & Mid(CStr(inDate),5,2) & Mid(CStr(inDate),1,2) & Mid(CStr(inDate),3,2)
								ConvertDate = CStr(outDate)
							Else
								outDate = "20" & Mid(CStr(inDate),5,2) & Mid(CStr(inDate),1,2) & Mid(CStr(inDate),3,2)
								ConvertDate = CStr(outDate)
							End If
						Case "YYMMDD":
							outDate = Mid(CStr(inDate),5,2) & Mid(CStr(inDate),1,2) & Mid(CStr(inDate),3,2)
							ConvertDate = CStr(outDate)
						Case "MMDDYY":
							ConvertDate = CStr(inDate)
						Case Else
							ConvertDate = ""
					End Select
				Else
					ConvertDate = ""
				End If
			Case "MMDDCCYY":
				If Len(CStr(inDate)) = 7 Then inDate = "0" & CStr(inDate)
				If IsDate(Mid(CStr(inDate),1,2) & "-" & Mid(CStr(inDate),3,2) & "-" & Mid(CStr(inDate),5,4)) Then
					Select Case outFormat
						Case "PC":
							outDate = Mid(CStr(inDate),1,2) & "-" & Mid(CStr(inDate),3,2) & "-" & Mid(CStr(inDate),5,4)
							ConvertDate = CDate(outDate)
						Case "CCYYMMDD":
							outDate = Mid(CStr(inDate),5,4) & Mid(CStr(inDate),1,2) & Mid(CStr(inDate),3,2)
							ConvertDate = CStr(outDate)
						Case "YYMMDD":
							outDate = Mid(CStr(inDate),7,2) & Mid(CStr(inDate),1,2) & Mid(CStr(inDate),3,2)
							ConvertDate = CStr(outDate)
						Case "MMDDYY":
							outDate = Mid(CStr(inDate),1,2) & Mid(CStr(inDate),3,2) & Mid(CStr(inDate),7,2)
							ConvertDate = CStr(outDate)
						Case "MMDDCCYY":
							ConvertDate = CStr(inDate)
						Case Else
							ConvertDate = ""
					End Select
				Else
					ConvertDate = ""
				End If
			Case "MMYY":
				If Len(CStr(inDate)) = 3 Then inDate = "0" & CStr(inDate)
				If IsDate(Mid(CStr(inDate),1,2) & "-1-" & Mid(CStr(inDate),3,2)) Then
					Select Case outFormat
						Case "PC":
							outDate = Mid(CStr(inDate),1,2) & "-1-" & Mid(CStr(inDate),3,2)
							ConvertDate = CDate(outDate)
						Case "CCYYMMDD":
							If Mid(Cstr(inDate),3,2) > 50 Then
								outDate = "19" & Mid(CStr(inDate),3,2) & Mid(CStr(inDate),1,2) & "01"
								ConvertDate = CStr(outDate)
							Else
								outDate = "20" & Mid(CStr(inDate),3,2) & Mid(CStr(inDate),1,2) & "01"
								ConvertDate = CStr(outDate)
							End If
						Case "YYMMDD":
							outDate = Mid(CStr(inDate),3,2) & Mid(CStr(inDate),1,2) & "01"
							ConvertDate = CStr(outDate)
						Case "MMDDYY":
							outDate = Mid(CStr(inDate),3,2) & "01" & Mid(CStr(inDate),1,2)
							ConvertDate = CStr(outDate)
						Case Else
							ConvertDate = ""
					End Select
				Else
					ConvertDate = ""
				End If
			Case "YY":
				If IsNumeric(inDate) Then
					Select Case outFormat
						Case "YYYY":
							If inDate < 10 Then
								ConvertDate = CInt("200" & CStr(inDate))
							ElseIf inDate < 50 Then
								ConvertDate = CInt("20" & CStr(inDate))
							Else
								ConvertDate = CInt("19" & CStr(inDate))
							End If
						Case Else
							ConvertDate = ""
					End Select
				Else
					ConvertDate = ""
				End If
			Case "PCTIME"
				If IsDate(inDate) Then
					Select Case outFormat
						Case "PCTIME":
							ConvertDate = CDate(inDate)
						Case "HH:MM:SS"
							If CInt(CStr(Hour(inDate))) < 10 Then
								outDate = "0" & CStr(Hour(inDate))
							Else
								outDate = CStr(Hour(inDate))
							End If
							If CInt(CStr(Minute(inDate))) < 10 Then
								outDate = outDate & ":0" & CStr(Minute(inDate))
							Else
								outDate = outDate & ":" & CStr(Minute(inDate))
							End If
							If CInt(CStr(Second(inDate))) < 10 Then
								outDate = outDate & ":0" & CStr(Second(inDate))
							Else
								outDate = outDate & ":" & CStr(Second(inDate))
							End If
							ConvertDate = outDate
						Case "HHMM"
							If CInt(CStr(Hour(inDate))) < 10 Then
								outDate = "0" & CStr(Hour(inDate))
							Else
								outDate = CStr(Hour(inDate))
							End If
							If CInt(CStr(Minute(inDate))) < 10 Then
								outDate = outDate & "0" & CStr(Minute(inDate))
							Else
								outDate = outDate & CStr(Minute(inDate))
							End If
							ConvertDate = outDate
						Case "HHMMSS"
							If CInt(CStr(Hour(inDate))) < 10 Then
								outDate = "0" & CStr(Hour(inDate))
							Else
								outDate = CStr(Hour(inDate))
							End If
							If CInt(CStr(Minute(inDate))) < 10 Then
								outDate = outDate & "0" & CStr(Minute(inDate))
							Else
								outDate = outDate & CStr(Minute(inDate))
							End If
							If CInt(CStr(Second(inDate))) < 10 Then
								outDate = outDate & "0" & CStr(Second(inDate))
							Else
								outDate = outDate & CStr(Second(inDate))
							End If
							ConvertDate = outDate
						Case Else
							ConvertDate = ""
					End Select
				End If
			Case "HHMMSS":
				If Len(CStr(inDate)) = 5 Then inDate = "0" & CStr(inDate)
				If Len(inDate) = 6 Then
					Select Case outFormat
						Case "PC":
							outDate = Mid(CStr(inDate),1,2) & ":" & Mid(CStr(inDate),3,2) & ":" & Mid(CStr(inDate),5,2)
							ConvertDate = CDate(outDate)
						Case Else
							ConvertDate = ""
					End Select
				Else
					ConvertDate = ""
				End If
			Case Else
				ConvertDate = ""
		End Select
	Else
		ConvertDate = ""
	End If
End Function
%>