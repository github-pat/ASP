Function Base64EncodeVB(inData)
	  Base64EncodeVB = Base64Encode(BinaryToString(inData))
	End Function

	Function Base64Encode(inData)
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
