Define Class XVFP as Session olepublic  
	
	Procedure run as String
		Lparameters cFileName As String, cOptions As String

		Local cReturn, cPath, cProgram, cFileLog, i, nLines, uReturnAux  
		
		If Empty(cOptions)
			cOptions = ""
		EndIf 
		
		cReturn = ""
		uReturnAux = ""
		
		Try 
			Dimension aLog[1]
			cFileLog = Addbs(Sys(2023)) + ForceExt(Sys(2015), "txt")
			
			SET ALTERNATE TO (cFileLog) ADDITIVE 
			SET ALTERNATE ON
		
			Compile (cFileName)
			cPath = JustPath(cFileName)	
			cProgram = JustStem(cFileName)
			Set path To (cPath) ADDITIVE 
			uReturnAux = Evaluate(cProgram+"()")
			
			SET ALTERNATE OFF
			SET ALTERNATE TO			
			
			nLines = ALines(aLog, FileToStr(cFileLog), 4)
			uReturnAux = Transform(uReturnAux)
			
			For i=1 To nLines
				cReturn = cReturn + Chrtran(aLog[i], Chr(26), "") + Chr(13) + Chr(10)
			Next 
			
			cReturn = cReturn + uReturnAux
			
			Delete File (cFileLog)			
			
			Do Case
				Case "browse" $ cOptions
					oBuild = CreateObject("CursorToString")
					cReturn = cReturn + Chr(13) + Chr(10) + oBuild.Build(Alias())
			EndCase 
						
		Catch To oErr
			cReturn = oErr.Message + " Linha: " + Transform(oErr.Lineno)
		EndTry 		
		
		Return cReturn
	EndProc
EndDefine 

Define Class CursorToString As Custom
	
	#DEFINE CharHorizontal "-"
	#DEFINE CharVertical "|"
	#DEFINE Border "+"
	
	Procedure Build
		Lparameters cCursor As String
		
		Local cCursorString, aStruct[1] 
		
		Afields(aStruct, cCursor)
		cCursorString = This.getHorizontalLines(@aStruct) + Chr(13) + Chr(10)
		cCursorString = cCursorString + This.getCaptionHeader(@aStruct) + Chr(13) + Chr(10)
		cCursorString = cCursorString + This.getHorizontalLines(@aStruct) + Chr(13) + Chr(10)
		
		Select (cCursor)
		Scan 
			cCursorString = cCursorString + This.getCellsValue(@aStruct) + Chr(13) + Chr(10)
		EndScan 

		cCursorString = cCursorString  + This.getHorizontalLines(@aStruct) + Chr(13) + Chr(10)

		Return cCursorString 		
		
	EndProc 
	
	Protected Procedure getHorizontalLines(aStruct)
		Local cHorizontalLines, cFieldName, nLen, nFields, i
		cHorizontalLines = ""
		nFields = Alen(aStruct, 1)
		For i=1 To nFields
			cFieldName = aStruct[i, 1]
			If aStruct[i, 2] = "M"
				nLen = 30
			Else 
				nLen = aStruct[i, 3]
			EndIf 
			If nLen < Len(cFieldName)
				nLen = Len(cFieldName)
			Endif

			cHorizontalLines = cHorizontalLines + Border + Replicate(CharHorizontal, nLen)
			
			If i=nFields
				cHorizontalLines = cHorizontalLines + Border
			EndIf 
		Next

		Return cHorizontalLines
	Endproc

	Protected Procedure getCaptionHeader(aStruct)
		Local cVerticalLines, cFieldName, nLen, nFields
		cVerticalLines = ""
		nFields = Alen(aStruct, 1)
		For i=1 To nFields
			cFieldName = aStruct[i, 1]
	 
			If aStruct[i, 2] = "M"
				nLen = 30
			Else 
				nLen = aStruct[i, 3]
			EndIf 
			
			If nLen < Len(cFieldName)
				nLen = Len(cFieldName)
			Endif
			
			lAddEndChar = i=nFields
			cVerticalLines = cVerticalLines + This.WriteValue(cFieldName, nLen, lAddEndChar)
		Next
		
		Return cVerticalLines
	EndProc

	Protected Procedure getCellsValue(aStruct)
		Local cVerticalLines, cFieldName, nLen, nFields, lAddEndChar
		cVerticalLines = ""
		nFields = Alen(aStruct, 1)
		For i=1 To nFields
			cFieldName = aStruct[i, 1]
			
			If aStruct[i, 2] = "M"
				nLen = 30
			Else 
				nLen = aStruct[i, 3]
			EndIf 	

			If nLen < Len(cFieldName)
				nLen = Len(cFieldName)
			EndIf
			
			cValue = Substr(Mline(Transform(&cFieldName), 1), 1, nLen)
			
			lAddEndChar = i=nFields

			cVerticalLines = cVerticalLines + This.WriteValue(cValue, nLen, lAddEndChar)
		Next
		
		Return cVerticalLines
	EndProc

	Protected Procedure WriteValue(cValue, nLen, lAddEndChar)
		Local cText
		cText = ""
		cText = cText + CharVertical + Transform(cValue) + Replicate(" ", Abs(nLen - Len(cValue)))
		
		If lAddEndChar
			cText = cText + CharVertical
		EndIf  
		
		Return cText
	Endproc
	
	
EndDefine 


