Function Base64Encode(inData)
	inData = BinaryToString(inData)
	Const Base64 = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
	Dim cOut, sOut, I
	  'For each group of 3 bytes
	For I = 1 To Len(inData) Step 3
	Dim nGroup, pOut, sGroup
	    'Create one long from this 3 bytes.
	    nGroup = &H10000 * Asc(Mid(inData, I, 1)) + _
	      &H100 * MyASC(Mid(inData, I + 1, 1)) + MyASC(Mid(inData, I + 2, 1))
	    'Oct splits the long To 8 groups with 3 bits
	    nGroup = Oct(nGroup)
	    'Add leading zeros
	    nGroup = String(8 - Len(nGroup), "0") & nGroup
	    'Convert To base64
	    pOut = Mid(Base64, CLng("&o" & Mid(nGroup, 1, 2)) + 1, 1) + _
	      Mid(Base64, CLng("&o" & Mid(nGroup, 3, 2)) + 1, 1) + _
	      Mid(Base64, CLng("&o" & Mid(nGroup, 5, 2)) + 1, 1) + _
	      Mid(Base64, CLng("&o" & Mid(nGroup, 7, 2)) + 1, 1)
	    sOut = sOut + pOut
	Next
	Select Case Len(inData) Mod 3
	    Case 1: '8 bit final
	      sOut = Left(sOut, Len(sOut) - 2) + "=="
	    Case 2: '16 bit final
	      sOut = Left(sOut, Len(sOut) - 1) + "="
	End Select
	Base64Encode = sOut
End Function

Function MyASC(OneChar)
  If OneChar = "" Then MyASC = 0 Else MyASC = Asc(OneChar)
End Function

function ValidaExtencion(ext)
	Select Case LCase(ext)
	  Case ".jpg"
		tipo = "image/jpg"
	  Case ".jpeg"
		tipo = "image/jpeg"
	  Case ".png"
		tipo = "image/png"
	  Case ".gif"
		tipo = "image/gif"
	  Case ".bmp"
		tipo = "image/bmp"
	  Case else
		tipo = null
	End Select 
end function
function XML()
	xml = "<dataset>"
	xml = xml & "<Plataforma>0</Plataforma>"
	xml = xml & "<CodigoSistema>1</CodigoSistema>"
	xml = xml & "<AreaDesarrollo>2</AreaDesarrollo>"
	xml = xml & "<TipoDocumento>99</TipoDocumento>"
	xml = xml & "<NameAPI>Cas_Service</NameAPI>"
	xml = xml & "<KeyAPI>39fafb5d3c1b1649e501cb87e522e9d3</KeyAPI>"
	xml = xml & "<User_Rut>"&rut&"</User_Rut>"
	xml = xml & "<Img_Firma>"&base64&"</Img_Firma>"
	xml = xml & "<Img_Nombre>"&name&"</Img_Nombre>"
	xml = xml & "<Img_Tipo>"&tipo&"</Img_Tipo>"
	xml = xml & "<Img_Tamano>"&tamano&"</Img_Tamano>"
	xml = xml & "</dataset>"
	response.write Certificado
end function
  Function BinaryToString(Binary)
    '2001 Antonin Foller, PSTRUH Software
    'Optimized version of PureASP conversion function
    'Selects the best algorithm to convert binary data to String data
    Dim TempString 

    On Error Resume Next
    'Recordset conversion has a best functionality
    TempString = RSBinaryToString(Binary)
    If Len(TempString) <> LenB(Binary) then'Conversion error
      'We have to use multibyte version of BinaryToString
      TempString = MBBinaryToString(Binary)
    end if
    BinaryToString = TempString
  End Function
