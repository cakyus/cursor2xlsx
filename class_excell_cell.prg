
DEFINE CLASS Excell_Cell AS Custom

	ColumnNumber = 0
	RowNumber = ''
	FoxproFieldType = ''
	Value = 0
	
	ZeroDateString = .F.
	
	FUNCTION GetFieldValueString

		lcFieldType = This.FoxproFieldType
		This.Value = This.Value
		
		DO CASE
			CASE lcFieldType = 'C'
				RETURN This.HTMLEncode(This.Value)
			CASE lcFieldType = 'N'
				RETURN LTRIM(STR(This.Value))
			CASE lcFieldType = 'D'
				IF This.ZeroDateString = .T. AND YEAR(This.Value) = 0 THEN
					RETURN '  -  -'
				ELSE
					RETURN LTRIM(STR(This.Value - DATE(1899,12,30)))
				ENDIF
			CASE lcFieldType = 'T'
			OTHERWISE
				ERROR 'Undefined foxpro field type ' + lcFieldType
		ENDCASE
		
		RETURN ''	
	ENDFUNC

	&& @link https://en.wikipedia.org/wiki/Character_encodings_in_HTML
	&& @notes Illegal characters
	&&	   HTML forbids the use of the characters with Universal Character Set/Unicode code points
	&&	       0 to 31, except 9, 10, and 13 (C0 control characters)
	&&	       127 (DEL character)
	&&	       128 to 159 (x80 – x9F, C1 control characters)
	&&	       55296 to 57343 (xD800 – xDFFF, the UTF-16 surrogate halves)
	&&     Numeric character reference
	&&     A numeric character reference in HTML refers to 
	&&     a character by its Universal Character Set/Unicode code point, and uses the format
	&&     &#nnnn; or &#xhhhh; 
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
			DO CASE
				CASE iCharAsc < 32 AND NOT INLIST(iCharAsc, 9, 10, 13)
					sChar = '&#' + PADL(iCharAsc, 4, '0') + ';'
				CASE iCharAsc = 127
					sChar = '&#' + PADL(iCharAsc, 4, '0') + ';'
				CASE iCharAsc > 127 AND iCharAsc < 160
					sChar = '&#' + PADL(iCharAsc, 4, '0') + ';'
				CASE iCharAsc > 55295 AND iCharAsc < 57344
					sChar = '&#' + PADL(iCharAsc, 4, '0') + ';'
				CASE sChar = '&'
					sChar = '&amp;'
				CASE sChar = '<'
					sChar = '&lt;'
				CASE sChar = '>'
					sChar = '&gt;'
				CASE sChar = '"'
					sChar = '&quote;'
				CASE sChar = "'"
					sChar = '&apos;'
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
