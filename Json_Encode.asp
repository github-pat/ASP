' Ejemplo de uso:
'	if not rs.eof then
'		data =  Json_Encode(rs.Fields, rs.GetRows())
'		response.write Json_Response(Array("Code", "Message", "Data"), Array(200, "Ok", data))
'	else
'		response.write Json_Response(Array("Code", "Message"), Array(404, "No se encontraron datos."))
'	end if

Function Json_Encode(key, value)
	rowLen = ubound(value, 2)
	x = "["
	y = "]"
	For i = 0 To rowLen
	 	j = 0
	 	x = x & "{"
		For each column in key
		 	x = x & """"& column.Name &""":"""& value(j,i) &""","
		 	j = j + 1
		Next
		x = left(x,len(x)-1) & "},"
	next
	Json_Encode = left(x,len(x)-1) & y
End Function
Function Json_Response(key, value)
	keyLen = ubound(key)
	x = "{"
	For i = 0 To keyLen
		if left(value(i), 1) = "{" or left(value(i), 1) = "[" then
			x = x & """"& key(i) &""":"& value(i) &","
		else
			x = x & """"& key(i) &""":"""& value(i) &""","
		end if 
	next
	Json_Response = left(x,len(x)-1) & "}"
End Function
