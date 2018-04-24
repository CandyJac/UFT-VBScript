'Call GenerateIdNonRSA("12","12","1993","Female","Namibia")
'Call GenerateIdRSA("15","01","1995","Male","RSA")
Public Function GenerateIdRSA(iDOBday,iDOBmth,iDOByear,sGender,sCitizenship)
	
	'Required Variables
	Dim iGender
	Dim iCitizenship
	Dim sID12
	Dim iSumOdd
	Dim sEvenDigits
	Dim iEvenDigits
	Dim iEvenDigitsX2 
	Dim sEvenDigitsX2 
	Dim sEvenX2_1
	Dim sEvenX2_2
	Dim sEvenX2_3
	Dim sEvenX2_4
	Dim sEvenX2_5
	Dim sEvenX2_6
	Dim sEvenX2_7
	Dim sEvenX2_8
	Dim sEvenX2_9
	Dim sEvenX2_10
	Dim iSumEvenX2
	Dim iSumOddEvenX2
	Dim sResult2nd
	Dim IResult2nd
	Dim iIDcd
	Dim sID13
	
	'GENERATE DIGITS
	'Gender
	If sGender = "Female" Then iGender = GetRandomNumber(0,4,0)
	If sGender = "Male" Then iGender = GetRandomNumber(5,9,0)
	
	'Citizenship
	If sCitizenship = "RSA" Then 	
		'South African
		iCitizenship = 1 	
	Else
		'Non South African
		iCitizenship = 0
	End If
		
	'1st 12 digits
	sID12 = Right(iDOByear,2) & Right(("0" & iDOBmth),2) & Right("0" & iDOBday,2) & CStr(iGender) & Right("000" & CStr(GetRandomNumber(0,999,0)),3) & CStr(iCitizenship) & Right("000" & CStr(GetRandomNumber(8,9,0)),1)
		
	'Sum of odd digits	
	iSumOdd = CInt(0 & Left(sID12,1)) + CInt(0 & Mid(sID12,3,1)) + CInt(0 & Mid(sID12,5,1)) + CInt(0 & Mid(sID12,7,1)) + CInt(0 & Mid(sID12,9,1)) + CInt(0 & Mid(sID12,11,1))
		
	'Even digits multiplied by two	
	sEvenDigits = "" & (Mid(sID12,2,1)) & (Mid(sID12,4,1)) & (Mid(sID12,6,1)) & (Mid(sID12,8,1)) & (Mid(sID12,10,1)) & (Right(sID12,1)) & ""
	iEvenDigits = Int(sEvenDigits)
	iEvenDigitsX2 = iEvenDigits * 2
	sEvenDigitsX2 = CStr(iEvenDigitsX2)
			
	'Extract dgits of sum of even digits multiplied by two	
	If NOT Left(sEvenDigitsX2,1) = "" Then
		sEvenX2_1 = Int(Left(sEvenDigitsX2,1))
	Else
		sEvenX2_1 = 0
	End If
	
	If NOT Mid(sEvenDigitsX2,2,1) = "" Then
		sEvenX2_2 = Int(Mid(sEvenDigitsX2,2,1))
	Else
		sEvenX2_2 = 0
	End If
	
	If NOT Mid(sEvenDigitsX2,3,1) = "" Then
		sEvenX2_3 = Int(Mid(sEvenDigitsX2,3,1))
	Else
		sEvenX2_3 = 0
	End If
	
	If NOT Mid(sEvenDigitsX2,4,1) = "" Then
		sEvenX2_4 = Int(Mid(sEvenDigitsX2,4,1))
	Else
		sEvenX2_4 = 0
	End If
	
	If NOT Mid(sEvenDigitsX2,5,1) = "" Then
		sEvenX2_5 = Int(Mid(sEvenDigitsX2,5,1))
	Else
		sEvenX2_5 = 0
	End If
	
	If NOT Mid(sEvenDigitsX2,6,1) = "" Then
		sEvenX2_6 = Int(Mid(sEvenDigitsX2,6,1))
	Else
		sEvenX2_6 = 0
	End If
	
	If NOT Mid(sEvenDigitsX2,7,1) = "" Then
		sEvenX2_7 = Int(Mid(sEvenDigitsX2,7,1))
	Else
		sEvenX2_7 = 0
	End If
	
	If NOT Mid(sEvenDigitsX2,8,1) = "" Then
		sEvenX2_8 = Int(Mid(sEvenDigitsX2,8,1))
	Else
		sEvenX2_8 = 0
	End If
	
	If NOT Mid(sEvenDigitsX2,9,1) = "" Then
		sEvenX2_9 = Int(Mid(sEvenDigitsX2,9,1))
	Else
		sEvenX2_9 = 0
	End If
	
	If NOT Mid(sEvenDigitsX2,10,1) = "" Then
		sEvenX2_10 = Int(Mid(sEvenDigitsX2,9,1))
	Else
		sEvenX2_10 = 0
	End If
	
	'Sum of the multiplied even digits	
	iSumEvenX2 = sEvenX2_1 + sEvenX2_2 + sEvenX2_3 + sEvenX2_4 + sEvenX2_5 + sEvenX2_6 + sEvenX2_7 + sEvenX2_8 + sEvenX2_9 + sEvenX2_10
	
	'Sum of the multiplied even digits plus sum of the odd digits	
	iSumOddEvenX2 = iSumEvenX2 + iSumOdd
		
	'Extract 2nd digit of result
	sResult2nd = Right(Left(iSumOddEvenX2,2),1)
	iResult2nd = CInt(sResult2nd)
		
	'Control Digit
	iIDcd = 10 - iResult2nd
	
	If iIDcd = 10 Then
		
		iIDcd = 0
		
	End If
		
	'Check for errors
	sID13 = sID12 & iIDcd
	If NOT Len(sID13) = 13 Then
	
		msgbox("Generated ID NOT 13 characters - check print log")
		Print("")
		print("Length sID13 = [ " & Len(sID13) & " ]")
		print("sID12 = [ " & sID12 & " ]")
		print("Length sID12 = [ " & Len(sID12) & " ]")
		print ("")
		print "d1 = " & CInt(Left(sID12,1)) 
		print "d2 = " & CInt(Mid(sID12,2,1))
		print "d3 = " & CInt(Mid(sID12,3,1))
		print "d4 = " & CInt(Mid(sID12,4,1))
		print "d5 = " & CInt(Mid(sID12,5,1))
		print "d6 = " & CInt(Mid(sID12,6,1))
		print "d7 = " & CInt(Mid(sID12,7,1))
		print "d8 = " & CInt(Mid(sID12,8,1))
		print "d9 = " & CInt(Mid(sID12,9,1))
		print "d10 = " & CInt(Mid(sID12,10,1))
		print "d11 = " & CInt(Mid(sID12,11,1))
		print "d12 = " & CInt(Right(sID12,1))
		print("")
		print("sResult2nd = [ " & sResult2nd & " ]")
		print("iIDcd = [ " & iIDcd & " ]")
	
	End If
	
	'Keep as string to keep trailing zeros for DOB 2000 - 2009
	GenerateIdRSA = "" & sID12 & iIDcd
	'MsgBox GenerateIdRSA
	
End Function

Public Function GetRandomNumber(iMin,iMax,iDecimalPlaces)

	Randomize		
	GetRandomNumber = Round(((Rnd * (iMax - iMin)) + iMin),iDecimalPlaces)
	
End Function

Public Function GenerateIdNonRSA(iDOBday,iDOBmth,iDOByear,sGender,sCountryOfIssue)
	
	Dim sGenderCharacter
	
	Select Case sCountryOfIssue
				
		Case "Mexico"
			
			If sGender = "Male" Then
				
				sGenderCharacter = "H"
				
			Else 

				sGenderCharacter = "M"
	
			End If
			
			GenerateIdNonRSA =  Right("0000" & GetRandomNumber(0,9999,0),4) & Right("00" & iDOByear,2) & Right("00" & iDOBmth,2) & Right("00" & iDOBday,2) & sGenderCharacter & Right(("0000000" & GetRandomNumber(1,9999999,0)),7)
			
			
		Case "Namibia"
		
			GenerateIdNonRSA = Right("00" & iDOByear,2) & Right("00" & iDOBmth,2) & Right("00" & iDOBday,2) & Right(("00000" & GetRandomNumber(1,99999,0)),5)
			
	
		Case "Swaziland"
			
			'08 or 09 possible that male and female (OR random)
			If sGender = "Male" Then
				
				sGenderCharacter = "08"
				
			Else 

				sGenderCharacter = "09"
	
			End If
				
			GenerateIdNonRSA = Right("00" & iDOByear,2) & Right("00" & iDOBmth,2) & Right("00" & iDOBday,2) & Right(("0000" & GetRandomNumber(1,9999,0)),4)  & sGenderCharacter & Right(("0" & GetRandomNumber(1,9,0)),1)
			
		
		Case "United States of America"
			
			'328-62-0845
			'023-28-7421
			'RANDOM|I-3|C--|I-2|C--|I-4
			
			GenerateIdNonRSA = Right(("00" & GetRandomNumber(1,999,0)),3) & "-" & Right(("0" & GetRandomNumber(1,99,0)),2) & "-" & Right(("000" & GetRandomNumber(1,9999,0)),4)
			
		Case "Zimbabwe"
		
			'1521439522673086195M
			'RANDOM|I-19|C-M
			
			If sGender = "Male" Then
				
				sGenderCharacter = "M"
				
			Else 

				sGenderCharacter = "F"
	
			End If
			
			GenerateIdNonRSA = Right(("000000000000000000" & GetRandomNumber(1,9999999999999999999,0)),19) & sGenderCharacter
	
		Case Else
		
			msgbox("Script update required for Non RSA ID generation - No Match : Country of Issue = [ " & sCountryOfIssue & "]")
			Print("Script update required for Non RSA ID generation - No Match : Country of Issue = [ " & sCountryOfIssue & "]") 
	
	End Select
	MsgBox GenerateIdNonRSA
	
End Function