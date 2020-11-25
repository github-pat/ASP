<!-- #include file="Inc/upload.inc.asp" -->
<%

'''
''' Base64Encode y Base64Decode son las funciones principales para codificar o descifrar un base64
'''
	Function Base64Decode(ByVal vCode)
	    Dim oXML, oNode
	    Set oXML = CreateObject("Msxml2.DOMDocument.3.0")
	    Set oNode = oXML.CreateElement("base64")
	    oNode.dataType = "bin.base64"
	    oNode.text = vCode
	    Base64Decode = Stream_BinaryToString(oNode.nodeTypedValue)
	    Set oNode = Nothing
	    Set oXML = Nothing
	End Function
	Function Stream_BinaryToString(Binary)
	  Const adTypeText = 2
	  Const adTypeBinary = 1
	  Dim BinaryStream 'As New Stream
	  Set BinaryStream = CreateObject("ADODB.Stream")
	  BinaryStream.Type = adTypeBinary
	  BinaryStream.Open
	  BinaryStream.Write Binary
	  BinaryStream.Position = 0
	  BinaryStream.Type = adTypeText
	  BinaryStream.CharSet = "iso-8859-1"
	  Stream_BinaryToString = BinaryStream.ReadText
	  Set BinaryStream = Nothing
	End Function
	Function Base64Encode(str)
		str = BinaryToString(str)
		Const Base64 = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
		Dim cOut, sOut, I
		For I = 1 To Len(str) Step 3
			Dim nGroup, pOut, sGroup
			nGroup = &H10000 * Asc(Mid(str, I, 1)) + _
			  &H100 * MyASC(Mid(str, I + 1, 1)) + MyASC(Mid(str, I + 2, 1))
			nGroup = Oct(nGroup)
			nGroup = String(8 - Len(nGroup), "0") & nGroup
			pOut = Mid(Base64, CLng("&o" & Mid(nGroup, 1, 2)) + 1, 1) + _
			  Mid(Base64, CLng("&o" & Mid(nGroup, 3, 2)) + 1, 1) + _
			  Mid(Base64, CLng("&o" & Mid(nGroup, 5, 2)) + 1, 1) + _
			  Mid(Base64, CLng("&o" & Mid(nGroup, 7, 2)) + 1, 1)
			sOut = sOut + pOut
		Next
		Select Case Len(str) Mod 3
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
	Function BinaryToString(Binary)
	    Dim TempString 
	    On Error Resume Next
	    TempString = RSBinaryToString(Binary)
	    If Len(TempString) <> LenB(Binary) then
	      TempString = MBBinaryToString(Binary)
	    end if
	    BinaryToString = TempString
	End Function

	Set Upload = GetUpload()
	file = Upload("txtFile").Value 'Archivo recibido de la vista, donde su name es txtFile
	base64 = Base64Encode(file)
	
	Response.Write base64
%>
