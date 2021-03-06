
DEFINE CLASS Excell_Cell AS Custom

	ColumnNumber = 0
	RowNumber = ''
	FoxproFieldType = ''
	Value = 0
	
	ZeroDateString = .F.
	
	&& Errors
	
	ErrorCount = 0
	ErrorLogPath = ''
	
	ErrorXMLChars = .F.
	ErrorXMLCharsText = 'Terdapat karakter yang tidak di-support format XML';
		+CHR(10)+CHR(13)+'REF: http://www.w3.org/TR/REC-xml/#charsets'
	
	ErrorExtendedAscii = .F.
	ErrorExtendedAsciiText = 'Terdapat karakter Extended ASCII';
		+CHR(10)+CHR(13)+'REF: http://docs.oracle.com/cd/E25054_01/relnotes.1111/e17596/toc.htm#sthref5'
		
	&& Menampilkan Error kalo ada karakter di luar standard Xml
	ShowErrorXmlChars = .T.
	
	&& Menampilkan Error kalo ada karakter Entended Ascii
	ShowErrorExtendedAscii = .T.
	
	PROCEDURE Init
		This.ErrorLogPath = SYS(2023) + '\FOX' + SYS(3) + '.LOG'
	ENDPROC
	
	FUNCTION ErrorLog
		LPARAMETERS sText
		This.ErrorCount = This.ErrorCount + 1
		STRTOFILE(sText+CHR(13), This.ErrorLogPath, .T.)
	ENDFUNC
	
	FUNCTION ErrorShow
		IF This.ErrorCount = 0 THEN
			RETURN .F.
		ENDIF
		IF This.ErrorXMLChars AND This.ShowErrorXmlChars THEN
			MESSAGEBOX(This.ErrorXMLCharsText;
				+CHR(10)+CHR(13)+'LOG: '+This.ErrorLogPath)
		ENDIF
		IF This.ErrorExtendedAscii AND This.ShowErrorExtendedAscii THEN
			MESSAGEBOX(This.ErrorExtendedAsciiText;
				+CHR(10)+CHR(13)+'LOG: '+This.ErrorLogPath)
		ENDIF
	ENDFUNC
	
	FUNCTION GetFieldValueString

		lcFieldType = This.FoxproFieldType
		This.Value = This.Value
		
		DO CASE
			CASE lcFieldType = 'C'
				IF VARTYPE(This.Value) = 'X'
					&& .NULL.
					RETURN ''
				ELSE
					RETURN This.HTMLEncode(This.Value)
				ENDIF
			CASE lcFieldType = 'N'
				IF VARTYPE(This.Value) = 'X'
					&& .NULL.
					RETURN '0'
				ELSE
					RETURN LTRIM(STR(This.Value))
				ENDIF
			CASE lcFieldType = 'D'
				IF VARTYPE(This.Value) = 'X'
					&& .NULL.
					IF This.ZeroDateString = .T. THEN
						RETURN '  -  -'
					ELSE
						RETURN LTRIM(STR(0))
					ENDIF
				ELSE
					IF This.ZeroDateString = .T. AND YEAR(This.Value) = 0 THEN
						RETURN '  -  -'
					ELSE
						RETURN LTRIM(STR(This.Value - DATE(1899,12,30)))
					ENDIF
				ENDIF
			OTHERWISE
				ERROR 'Undefined foxpro field type ' + lcFieldType
		ENDCASE
		
		RETURN ''	
	ENDFUNC

	&& @link http://www.w3.org/TR/REC-xml/#charsets
	&& @notes character yang disupport oleh standard internasional:
	&&        9, 10, 13, 32-55295, 57344-65533, dan 65536-1114111


	FUNCTION IsXmlChars
		LPARAMETERS iCharAsc
		IF INLIST(iCharAsc, 9, 10, 13) ;
			OR BETWEEN(iCharAsc, 32, 55295) ;
			OR BETWEEN(iCharAsc, 57344, 65533) ;
			OR BETWEEN(iCharAsc, 65536, 1114111) ;
			THEN
			RETURN .T.
		ELSE
			RETURN .F.
		ENDIF		
	ENDFUNC
	
	&& @link http://msdn.microsoft.com/en-us/library/9hxt0028%28v=vs.80%29.aspx
	&& @notes ascii range 128-255
	
	FUNCTION IsExtendedAscii
		LPARAMETERS iCharAsc
		IF BETWEEN(iCharAsc, 128, 255) THEN
			RETURN .T.
		ELSE
			RETURN .F.
		ENDIF		
	ENDFUNC
	
	&& @link https://en.wikipedia.org/wiki/Character_encodings_in_HTML
	&& @notes XML character references
	&&     &amp;  & (ampersand, U+0026)
	&&     &lt;   < (less-than sign, U+003C)
	&&     &gt;   > (greater-than sign, U+003E)
	&&     &quot; " (quotation mark, U+0022)
	&&     &apos; ' (apostrophe, U+0027)
	
	FUNCTION HTMLEncode
		LPARAMETERS sText
		LOCAL i, j, sTextResult, sChar, iCharAsc
		
		sTextResult = ''
		sText = RTRIM(sText)
		FOR i = 1 TO LEN(sText)
			sChar = SUBSTR(sText, i, 1)
			iCharAsc = ASC(sChar)
			
			&& XML Supported characters
			IF This.IsXmlChars(iCharAsc) = .F. THEN
				This.ErrorXMLChars = .T.
				This.ErrorLog('Excell_Cell.HTMLEncode IsXmlChars ' + LTRIM(STR(iCharAsc)) + ' ' + sChar)
			ENDIF
			
			&& Extended ASCII Characters
			IF This.IsExtendedAscii(iCharAsc) = .T. THEN
				This.ErrorExtendedAscii = .T.
				This.ErrorLog('Excell_Cell.HTMLEncode IsExtendedAscii ' + LTRIM(STR(iCharAsc)) + ' ' + sChar)
			ENDIF
			
			DO CASE
				&& HTML Special Characters
				CASE sChar = '&'
					sChar = '&amp;'
				CASE sChar = '<'
					sChar = '&lt;'
				CASE sChar = '>'
					sChar = '&gt;'
				CASE sChar = '"'
					sChar = '&quot;'
				CASE sChar = "'"
					sChar = '&apos;'
				&& Space + Keyboard One-Stroke Characters
				CASE BETWEEN(iCharAsc, 32, 126)
					&& Not Encoded. For what 0 to 127 looks like,
					&& see http://msdn.microsoft.com/en-us/library/60ecse8t%28v=vs.80%29.aspx
				OTHERWISE
					sChar = '&#' + PADL(iCharAsc, 4, '0') + ';'				
			ENDCASE
			
			sTextResult = sTextResult + sChar
		ENDFOR
		
		RETURN sTextResult
	ENDFUNC
	
	FUNCTION GetColumnNameFromNumber
		LOCAL lnReminder, lcResult, lnLength
		
		lnNumber = This.ColumnNumber
		lcResult = ''
		lnLength = 1
		
		DO WHILE .T.
		 
			lnReminder = MOD(lnNumber , 26)
			lnResult = FLOOR(lnNumber / 26)

			IF lnReminder = 0
				lcResult = 'Z' + lcResult
			ELSE
				lcResult = CHR(64 + lnReminder) + lcResult
			ENDIF
			
			lnLength = lnLength + 1
			lnNumber = lnResult
			
			IF lnResult =< 26
				IF lnResult > 0
					IF lnReminder > 0
						lcResult = CHR(64 + lnResult) + lcResult
					ELSE
						IF lnResult > 1
							lcResult = CHR(64 + lnResult - 1) + lcResult
						ENDIF
					ENDIF
				ENDIF
				EXIT
			ENDIF
			
		ENDDO
		
		RETURN lcResult
	ENDFUNC
	
	FUNCTION GetExcellFieldType
	
&& lcFoxproFieldType
&& A Array (only returned when the optional 1 parameter is included)
&& C Character, Varchar, Varchar (Binary)
&& D Date
&& G General
&& L Logical
&& M Memo
&& N Numeric, Float, Double, or Integer
&& O Object
&& Q Blob, Varbinary
&& S Screen
&& T DateTime
&& U Undefined type of expression or cannot evaluate expression.
&& Y Currency

&& lcExcellCellDataType
&& const TYPE_STRING2		= 'str';
&& const TYPE_STRING		= 's';
&& const TYPE_FORMULA		= 'f';
&& const TYPE_NUMERIC		= 'n';
&& const TYPE_BOOL			= 'b';
&& const TYPE_NULL			= 's';
&& const TYPE_INLINE		= 'inlineStr';
&& const TYPE_ERROR		= 'e';

		DO CASE
			CASE This.FoxproFieldType = 'C'
				RETURN 't="str"'
			CASE This.FoxproFieldType = 'N'
				RETURN 't="n"'
			CASE This.FoxproFieldType = 'D'
				IF This.ZeroDateString = .T. AND YEAR(This.Value) = 0 THEN
					RETURN 't="str"'
				ELSE
					RETURN 's="1" t="n"'
				ENDIF
			CASE This.FoxproFieldType = 'T'
			OTHERWISE
				ERROR 'Undefined foxpro field type ' + This.FoxproFieldType
		ENDCASE
		RETURN 't="str"'
	ENDFUNC
	
	FUNCTION ToString
		LOCAL lcRange, lcType, lcValue
		
		lcRange = 'r="' ;
			+ This.GetColumnNameFromNumber() ;
			+ This.RowNumber ;
			+ '"'
		lcType = This.GetExcellFieldType()
		lcValue = This.GetFieldValueString()
		
		RETURN '<c ' + lcRange + ' ' + lcType + '><v>' + lcValue + '</v></c>'
	ENDFUNC
ENDDEFINE
