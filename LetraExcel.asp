	' Obtiene letra de columna para exportacion a excel
  Function LetraExcel(num)
		if num < 26 then
			letra = Chr(num+65)
		else
			letra =	Chr(Fix(num/26)+64) 'num=>26 = 1
			cont  =	Fix(num/26)
			letra = letra & Chr(Fix(num-(26*cont)+65))
		end if
		LetraExcel = letra
	End Function
