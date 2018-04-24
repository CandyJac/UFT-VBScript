Public Function DateOfBirth
	
Dim Day
Dim Max,Min
	Max=30
	Min=1
	Randomize
Day = (Int((max-min+1)*Rnd+min))

Dim Month
Dim Max1,Min1
	Max1=12
	Min1=1
	Randomize
Month = (Int((max1-min1+1)*Rnd+min1))

Dim Year
Dim Max2,Min2
	Max2=1995
	Min2=2000
	Randomize
Year = (Int((max2-min2+1)*Rnd+min2))

DateOfBirth = GenerateIdRSA(CStr(Day),CStr(Month),CStr(Year),"Male","RSA")

End Function