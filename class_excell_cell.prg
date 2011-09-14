
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
				RETURN RTRIM(This.Value)
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
