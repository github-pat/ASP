Function Base64Encode(str)
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

Function Base64Decode(ByVal base64String)
	Const Base64 = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
	Dim dataLength, sOut, groupBegin
	base64String = Replace(base64String, vbCrLf, "")
	base64String = Replace(base64String, vbTab, "")
	base64String = Replace(base64String, " ", "")
	dataLength = Len(base64String)
	If dataLength Mod 4 <> 0 Then
		Err.Raise 1, "Base64Decode", "Bad Base64 string."
	Exit Function
	End If
	For groupBegin = 1 To dataLength Step 4
		Dim numDataBytes, CharCounter, thisChar, thisData, nGroup, pOut
		numDataBytes = 3
		nGroup = 0
		For CharCounter = 0 To 3
			thisChar = Mid(base64String, groupBegin + CharCounter, 1)
			If thisChar = "=" Then
				numDataBytes = numDataBytes - 1
				thisData = 0
			Else
				thisData = InStr(1, Base64, thisChar, vbBinaryCompare) - 1
			End If
			If thisData = -1 Then
				Err.Raise 2, "Base64Decode", "Bad character In Base64 string."
			Exit Function
			End If
			nGroup = 64 * nGroup + thisData
		Next
		nGroup = Hex(nGroup)
		nGroup = String(6 - Len(nGroup), "0") & nGroup
		pOut = Chr(CByte("&H" & Mid(nGroup, 1, 2))) + _
		  Chr(CByte("&H" & Mid(nGroup, 3, 2))) + _
		  Chr(CByte("&H" & Mid(nGroup, 5, 2)))
		sOut = sOut & Left(pOut, numDataBytes)
	Next
	Base64Decode = sOut
End Function

' Para archivos, previo a codificar
Function BinaryToString(Binary)
    Dim TempString 
    On Error Resume Next
    TempString = RSBinaryToString(Binary)
    If Len(TempString) <> LenB(Binary) then
      TempString = MBBinaryToString(Binary)
    end if
    BinaryToString = TempString
End Function



'Acá la llamas
dim file = Upload("File").Value
base64 = Base64Encode(BinaryToString(file))
