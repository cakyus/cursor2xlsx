
DEFINE CLASS Excell_Cell AS Custom

	ColumnNumber = 0
	RowNumber = ''
	FoxproFieldType = ''
	Value = 0
	
	FUNCTION GetFieldValueString

		lcFieldType = This.FoxproFieldType
		laFieldValue = This.Value
		
		DO CASE
			CASE lcFieldType = 'C'
				RETURN RTRIM(laFieldValue)
			CASE lcFieldType = 'N'
				RETURN LTRIM(STR(laFieldValue))
			CASE lcFieldType = 'D'
				&& RETURN LTRIM(STR(laFieldValue - DATE(1900,1,1)))
			CASE lcFieldType = 'T'
			OTHERWISE
				ERROR 'Undefined foxpro field type ' + lcFieldType
		ENDCASE
		RETURN ''	
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
				RETURN 'str'
			CASE This.FoxproFieldType = 'N'
				RETURN 'n'
			CASE This.FoxproFieldType = 'D'
			CASE This.FoxproFieldType = 'T'
			OTHERWISE
				ERROR 'Undefined foxpro field type ' + This.FoxproFieldType
		ENDCASE
		RETURN 'str'	
	ENDFUNC
	
	FUNCTION ToString
		LOCAL lcRange, lcType, lcValue
		
		lcRange = 'r="' ;
			+ This.GetColumnNameFromNumber() ;
			+ This.RowNumber ;
			+ '"'
		lcType = 't="' + This.GetExcellFieldType() + '"'
		lcValue = This.GetFieldValueString()
		
		RETURN '<c ' + lcRange + ' ' + lcType + '><v>' + lcValue + '</v></c>'
	ENDFUNC
ENDDEFINE
